from datetime import datetime
from spyre.Models.sales_models import *
from spyre.Models.shared_models import *
import pandas as pd
from spyre.spire import Spire, SpireClient
import logging
from collections import defaultdict

#--------------------------------------------------------------------------------------------------------------------#

"""
This is a script that imports FBA Orders into Spire

-   To start, download a report from Amazon From the link below, using Dated Ranges and Order Date Type
    https://sellercentral.amazon.ca/reportcentral/FlatFileAllOrdersReport/1
-   Create an excel file from the downloaded text file
-   Filter out FBM Orders - Fullfilment Channel Column with values Merchant
-   Filter out Disposals - Sales Channel Column with values Non-Amazon
-   Filter out Cancelled Orders
-   Define EXCEL_FILE, SHEET_NAME, LOG_FILE Variables

"""

EXCEL_FILE = "FBA Orders August Amazon Report.xlsx"
SHEET_NAME = "FBA Orders August Amazon Report"
LOG_FILE = 'fba-orders-import-august-log'
SPIRE_USERNAME = ''
SPIRE_PASSWORD = ''
SPIRE_COMPANY = ''
SPIRE_HOST = ''


#--------------------------------------------------------------------------------------------------------------------#


TAXES = {
    0.05 : 1,
    0.08 : 2,
    0.13 : 3,
    0.15 : 4,
    0.14 : 5,
    0.12 : 6,
    0.11 : 7 
}

def setup_logger(name='my_logger', log_file=LOG_FILE, level=logging.INFO):
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.propagate = False

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)

    if not logger.handlers:
        logger.addHandler(console_handler)
        logger.addHandler(file_handler)

    return logger


def convert_order_date(date_str: str) -> str:
    for fmt in ("%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""
    

def get_tax_code(order: dict) -> int:
    try:
        item_price = float(order.get('item-price', 0))
        item_tax = float(order.get('item-tax', 0))

        if item_price == 0:
            raise ValueError("Item price cannot be zero.")

        tax_rate = item_tax / item_price
        # Find the closest tax rate key
        closest_rate = min(TAXES.keys(), key=lambda r: abs(r - tax_rate))
        return TAXES[closest_rate]

    except (TypeError, ValueError):
        return None  # or raise an error if preferred


def build_sales_orders(records: list[dict]) -> list[SalesOrder]:
    """
    Groups sales order records in pandas data frame by amazon id
    Converts a grouped record into a SalesOrder Object
    Returns a list of SalesOrder objects
    """

    orders_by_id = defaultdict(list)

    # Group rows by order-id
    for record in records:
        orders_by_id[record.get("amazon-order-id")].append(record)

    sales_orders = []

    for amazn_id, group in orders_by_id.items():
        valid_rows = [r for r in group if r.get("item-status") != "Cancelled"]
        if not valid_rows:
            continue  # skip whole order if all rows are cancelled

        first_row = valid_rows[0]  # Use first non-cancelled row for order-level info  # Common fields from first row
        
        if first_row.get('fulfillment-channel') == 'Merchant':
            continue
        if first_row.get('order-status') == 'cancelled':
            continue
        if first_row.get('sales-channel') == 'Non-Amazon':
            continue

        shipped = first_row.get('order-status') == 'Shipped'
        udf = {
            "shopid": amazn_id,
            "shipped": shipped
        }

        tax_code = get_tax_code(first_row)

        shipping_address = Address(
            city=first_row.get('ship-city'),
            provState=first_row.get('ship-state'),
            postalCode=first_row.get('ship-postal-code'),
            country=first_row.get('ship-country'),
            salesTaxes=[{"code": tax_code}]
        )

        # Build items list
        items = []
        for row in valid_rows:
            if row.get("item-status") == "Cancelled":
                continue
            item = SalesOrderItem(
                inventory=Inventory(partNo=row.get("sku"), whse='AMZN'),
                partNo=row.get("sku"),
                orderQty=str(row.get("quantity")),
                unitPrice=str(float(row.get("item-price")) / float(row.get("quantity"))),
                taxFlags=[True, True, True, True]
            )
            items.append(item)

        def safe_float(val):
            try:
                f = float(val)
                if f != f:  # NaN check
                    return 0.0
                return f
            except (ValueError, TypeError):
                return 0.0

        # Sum freight across rows
        freight = str(sum(
            safe_float(r.get('shipping-price')) - safe_float(r.get('ship-promotion-discount'))
            for r in valid_rows
        ))

        customer = Customer(customerNo="AMAZON")

        sales_orders.append(
            SalesOrder(
                orderDate=convert_order_date(first_row.get("purchase-date")),
                type="O",
                referenceNo=amazn_id,
                hold=not shipped,
                status="O",
                currency=Currency(code=first_row.get("currency")) if first_row.get("currency") else None,
                shippingAddress=shipping_address,
                items=items,
                udf=udf,
                freight=freight,
                customer=customer,
            )
        )

    return sales_orders

df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME , dtype=str)
df.fillna("0.0")
records = df.to_dict(orient='records')

logger = setup_logger()
errors = []

spire_client = SpireClient(company=SPIRE_COMPANY, username=SPIRE_USERNAME, password=SPIRE_PASSWORD, host=SPIRE_HOST, secure=False)
spire = Spire(client=spire_client)

sales_orders = build_sales_orders(records=records)

for sales_order in sales_orders:

    try:
        order = spire.orders.create_sales_order(sales_order)
    except Exception as e:
        logger.error(f"Error Creating Order {order.referenceNo} | Error : {e}")
        continue

    # Check for backorder, if backordered, skip payment and invoicing. Also skip for pending orders
    backordered = False
    for item in order.model.items:
        if int(item.backorderQty) > 0:
            backordered = True
    if backordered or order.model.hold:
        logger.info(f"Order No {order.referenceNo} Successfulyy Created, Skipped Invoicing | Backordered : {backordered} | Pending : {order.model.hold}")
        continue

    try:
        order.model.payments = [ { "method" : "06" , "amount" : order.model.total }]
        order.update()
    except Exception as e:
        logger.error(f"Error adding payment to Order {order.referenceNo} | Error : {e}")
        continue
    
    try:    
        order.invoice()
    except Exception as e:
        logger.error(f"Error invoicing Order {order.referenceNo} | Error : {e}")
        continue

    logger.info(f"Order No {order.referenceNo} Successfulyy Created and Invoiced")
