// =============================================================
// CRM Order Automation — Power BI M Queries
// =============================================================
// HOW TO USE:
//   1. Open Power BI Desktop
//   2. Home -> Transform Data -> Power Query Editor
//   3. Home -> New Source -> Blank Query
//   4. View -> Advanced Editor
//   5. Paste the query for each table below
//   6. Rename the query to match the table name
//   7. Close & Apply
// =============================================================


// -------------------------------------------------------------
// TABLE 1: Runs
// One row per automation session.
// Paste this as query named: Runs
// -------------------------------------------------------------
let
    Source       = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/runs")),
    Value        = Source[value],
    ToTable      = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded     = Table.ExpandRecordColumn(
                       ToTable, "Column1",
                       {"id","started_at","finished_at","excel_filename",
                        "status","error_message","total_documents",
                        "total_lines","report_path","customer_name","created_at"},
                       {"id","started_at","finished_at","excel_filename",
                        "status","error_message","total_documents",
                        "total_lines","report_path","customer_name","created_at"}
                   ),
    TypedTable   = Table.TransformColumnTypes(Expanded, {
                       {"id",               Int64.Type},
                       {"started_at",       type datetime},
                       {"finished_at",      type datetime},
                       {"excel_filename",   type text},
                       {"status",           type text},
                       {"error_message",    type text},
                       {"total_documents",  Int64.Type},
                       {"total_lines",      Int64.Type},
                       {"report_path",      type text},
                       {"customer_name",    type text},
                       {"created_at",       type datetime}
                   })
in
    TypedTable


// -------------------------------------------------------------
// TABLE 2: PostingRows
// Every CRM posting report row across all runs.
// Paste this as query named: PostingRows
// -------------------------------------------------------------
let
    Source       = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/posting_rows")),
    Value        = Source[value],
    ToTable      = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded     = Table.ExpandRecordColumn(
                       ToTable, "Column1",
                       {"id","run_id","posting_date","customer_no","document_no",
                        "excel_item_no","excel_description","excel_qty",
                        "excel_unit_price","crm_item_no","crm_product",
                        "crm_inventory","allocated_qty","crm_unit_price",
                        "inventory_status"},
                       {"id","run_id","posting_date","customer_no","document_no",
                        "excel_item_no","excel_description","excel_qty",
                        "excel_unit_price","crm_item_no","crm_product",
                        "crm_inventory","allocated_qty","crm_unit_price",
                        "inventory_status"}
                   ),
    TypedTable   = Table.TransformColumnTypes(Expanded, {
                       {"id",               Int64.Type},
                       {"run_id",           Int64.Type},
                       {"posting_date",     type date},
                       {"customer_no",      type text},
                       {"document_no",      type text},
                       {"excel_item_no",    type text},
                       {"excel_description",type text},
                       {"excel_qty",        type number},
                       {"excel_unit_price", type number},
                       {"crm_item_no",      type text},
                       {"crm_product",      type text},
                       {"crm_inventory",    type number},
                       {"allocated_qty",    type number},
                       {"crm_unit_price",   type number},
                       {"inventory_status", type text}
                   })
in
    TypedTable


// -------------------------------------------------------------
// TABLE 3: ExcelLines
// Every Excel order line loaded across all runs.
// Paste this as query named: ExcelLines
// -------------------------------------------------------------
let
    Source       = Json.Document(Web.Contents("http://127.0.0.1:8000/odata/excel_lines")),
    Value        = Source[value],
    ToTable      = Table.FromList(Value, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    Expanded     = Table.ExpandRecordColumn(
                       ToTable, "Column1",
                       {"id","run_id","document_no","customer_no","item_no",
                        "description","quantity","unit_price","amount",
                        "unit_of_measure","location_code","crm_search_text"},
                       {"id","run_id","document_no","customer_no","item_no",
                        "description","quantity","unit_price","amount",
                        "unit_of_measure","location_code","crm_search_text"}
                   ),
    TypedTable   = Table.TransformColumnTypes(Expanded, {
                       {"id",              Int64.Type},
                       {"run_id",          Int64.Type},
                       {"document_no",     type text},
                       {"customer_no",     type text},
                       {"item_no",         type text},
                       {"description",     type text},
                       {"quantity",        type number},
                       {"unit_price",      type number},
                       {"amount",          type number},
                       {"unit_of_measure", type text},
                       {"location_code",   type text},
                       {"crm_search_text", type text}
                   })
in
    TypedTable
