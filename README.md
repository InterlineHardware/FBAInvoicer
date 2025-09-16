# Amazon Invoicing Guide

## Step 1: Download Monthly Orders Report
- Go to: [Amazon All Orders Report](https://sellercentral.amazon.ca/reportcentral/FlatFileAllOrdersReport/1)
- Select **Date Range** and **Order Date Type**.
- Example: For August, pick **08/01 to 08/31**.
- Dowloand Report.

## Step 2: Create Master Monthly Report
1. Open a new Excel Workbook.
2. Go to **Data → From Text/CSV**.
3. Import the downloaded text file.
4. Save the workbook as **FBA Orders {Month} Amazon Report Master**.

## Step 3: Prepare FBA Orders File
1. Make a copy of the **Master Monthly Report** and open it.
2. Apply filters:
   - **Fulfillment Channel**: Remove `Merchant` (keep only FBA).
   - **Sales Channel**: Remove `Non-Amazon` (filter out disposals).
   - **Order Status**: Remove `Cancelled` (remove cancelled orders).

## Step 4: Prepare and Run Script
- Note: Latest version of spyreapi required. install with `pip install --upgrade spyreapi`
1. Define variables in your script:
   - `EXCEL_FILE`
   - `SHEET_NAME`
   - `LOG_FILE`
   - `SPIRE_USERNAME`
   - `SPIRE_PASSWORD`
   - `SPIRE_COMPANY`
   - `SPIRE_HOST #spirehost:port`
2. Run the script.
3. From the filedused for running the script, copy the records into a **Master FBA Orders File**.
   - This keeps track of all imported orders.
   - This file is also used for processing returns.

## Step 5: Download Returns Report
- Go to: [Amazon Returns Report](https://sellercentral.amazon.ca/reportcentral/CUSTOMER_RETURNS/1)
- Select **Date Range**:
  - For monthly import (e.g. August), pick **08/01 to today**.
- Note: Returns are dated by the **return request date**, and can include orders from prior months.

## Step 6: Process Returns Report
1. Download the returns CSV file.
2. Import into a new Excel Workbook.
3. Merge with the **Master FBA Orders Sheet**.
4. Delete any returns where there is **no corresponding order**.

---
