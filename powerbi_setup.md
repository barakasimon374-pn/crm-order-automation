# Power BI Integration Guide — CRM Order Automation

## Overview

The automation server exposes three OData-compatible JSON endpoints that
Power BI connects to directly using the built-in **Web** connector.
No custom connector or gateway is required while working locally.

---

## Prerequisites

- Power BI Desktop installed
- The automation server running at `http://127.0.0.1:8000`
  (run `python excel_upload_page.py` in the project folder)

---

## Step 1 — Connect Power BI to the data

### Option A — Paste M queries directly (recommended)

1. Open **Power BI Desktop**
2. Click **Home → Transform Data** to open Power Query Editor
3. Click **Home → New Source → Blank Query**
4. Click **View → Advanced Editor**
5. Delete the default text, paste the query from `powerbi_queries.m`
   for the table you want (Runs, PostingRows, or ExcelLines)
6. Click **Done**
7. Rename the query in the left panel to match the table name
8. Repeat steps 3–7 for each of the three tables
9. Click **Close & Apply**

### Option B — Use the Web connector

1. Click **Home → Get Data → Web**
2. Enter one of these URLs:
   - `http://127.0.0.1:8000/odata/runs`
   - `http://127.0.0.1:8000/odata/posting_rows`
   - `http://127.0.0.1:8000/odata/excel_lines`
3. Click **OK → Connect**
4. In the Navigator, select **value** from the JSON tree
5. Click **Transform Data**, then expand the Record column
6. Set column types as needed

---

## Step 2 — Create relationships

In **Model view**, create these relationships:

| From table  | Column | To table    | Column |
|-------------|--------|-------------|--------|
| PostingRows | run_id | Runs        | id     |
| ExcelLines  | run_id | Runs        | id     |

Cardinality: Many-to-One, Single direction.

---

## Step 3 — Suggested measures (DAX)

Paste these into **Home → New Measure** on the PostingRows table:

```dax
Total Lines = COUNTROWS(PostingRows)

Enough Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "ENOUGH INVENTORY")

Insufficient Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "INSUFFICIENT INVENTORY")

Zero Inventory =
    CALCULATE(COUNTROWS(PostingRows),
              PostingRows[inventory_status] = "ZERO INVENTORY")

Fulfilment Rate % =
    DIVIDE([Enough Inventory], [Total Lines], 0) * 100

Total Allocated Qty = SUM(PostingRows[allocated_qty])

Total Excel Qty = SUM(PostingRows[excel_qty])

Shortfall Qty = [Total Excel Qty] - [Total Allocated Qty]
```

---

## Step 4 — Suggested visuals

### Page 1 — Run Summary
- **Card** visuals: Total Lines, Fulfilment Rate %, Total Runs
- **Bar chart**: Runs[started_at] vs Total Lines (trend over time)
- **Pie/Donut**: Inventory status breakdown (Enough / Insufficient / Zero)
- **Table**: Runs table with status, customer_name, excel_filename, total_lines

### Page 2 — Posting Detail
- **Slicer**: Runs[started_at] date range
- **Slicer**: PostingRows[inventory_status]
- **Slicer**: PostingRows[customer_no]
- **Table**: PostingRows — document_no, excel_item_no, excel_description,
  excel_qty, crm_product, crm_inventory, allocated_qty, inventory_status
- **Conditional formatting** on inventory_status column:
  - ENOUGH INVENTORY → green
  - INSUFFICIENT INVENTORY → yellow
  - ZERO INVENTORY → red

### Page 3 — Product Analysis
- **Bar chart**: crm_item_no vs Total Allocated Qty (top products)
- **Scatter**: excel_qty (x) vs crm_inventory (y) coloured by inventory_status
- **Table**: ExcelLines grouped by item_no — count of orders, total quantity

---

## Step 5 — Refresh data

Click **Home → Refresh** in Power BI Desktop to pull the latest data
from the running server at any time.

For scheduled refresh via Power BI Service, you would need a
**Power BI Gateway** configured for the local server, or publish the
server to a reachable URL first.

---

## Available endpoints

| Endpoint | Description |
|---|---|
| `GET /odata/runs` | All automation runs |
| `GET /odata/posting_rows` | All posting report rows (optional `?run_id=N`) |
| `GET /odata/excel_lines` | All Excel order lines (optional `?run_id=N`) |
| `GET /odata` | Service document listing all endpoints |
