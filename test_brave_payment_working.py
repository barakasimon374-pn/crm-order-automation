from collections import Counter
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import traceback
import sys
import os
# Prevent Windows console encoding errors for Unicode CRM text.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(
        encoding="utf-8",
        errors="backslashreplace"
    )

if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(
        encoding="utf-8",
        errors="backslashreplace"
    )
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONFIGURATION
# ============================================================

CRM_URL = os.environ.get("CRM_URL", "").strip()

BRAVE_PATH = (
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
)

CUSTOMER_NAME_SEQUENCE = (
    "Aaron", "Alice",
    "Brian", "Brenda",
    "Charles", "Chloe",
    "Daniel", "Diana",
    "Edward", "Emily",
    "Frank", "Fiona",
    "George", "Grace",
    "Harry", "Hazel",
    "Ian", "Ivy",
    "Jack", "Jane",
    "Kevin", "Kylie",
    "Liam", "Lucy",
    "Mark", "Mary",
    "Nathan", "Nora",
    "Oliver", "Olivia",
    "Paul", "Paula",
    "Quentin", "Quinn",
    "Robert", "Rose",
    "Samuel", "Sarah",
    "Thomas", "Tina",
    "Ulysses", "Uma",
    "Victor", "Victoria",
    "Walter", "Wendy",
    "Xavier", "Xenia",
    "Yusuf", "Yvonne",
    "Zach", "Zoe",
)

CUSTOMER_NAME_STATE_FILE = (
    Path(__file__).resolve().parent
    / "customer_name_rotation.txt"
)
POSTING_REPORT_ROWS = []

CURRENT_REPORT_CONTEXT = {
    "customer_name": "",
    "document_no": "",
    "excel_item_no": "",
    "excel_description": "",
    "excel_quantity": 0,
}


WAREHOUSE_NAME = "SUAM STORES"
EXCEL_FILE = (
    Path(r"C:\Users\user\Downloads\Posted Sales Invoice Lines - 2026-09-26T080646.843.xlsx")
)

EXCEL_PREVIEW_ONLY = False

EXCLUDED_FT_DESCRIPTIONS = (
    "FT 5209 40.5*40.5 12PCS",
)
IGNORED_ITEM_CODES = {
    "24102",
    "33551",
    "33518",
    "33883",
}
def is_approved_excel_description(description):
    """
    Return True only for descriptions that are allowed
    to continue toward CRM processing.

    Approved:
        ASIAN TOILET...
        FRENCIA...
        FT...
        WT...

    FT exclusions:
        FT GG...
        FT GS...
        Explicit excluded FT descriptions.
    """

    text = str(description).strip()
    upper_text = text.upper()

    if "ASIAN TOILET" in upper_text:
        return True

    if upper_text.startswith("FRENCIA"):
        return True

    if upper_text.startswith("WT "):
        return True

    if upper_text.startswith("FT "):

        # Ignore FT products beginning with GG or GS.
        if (
            upper_text.startswith("FT GG")
            or upper_text.startswith("FT GS")
        ):
            return False

        normalized = " ".join(upper_text.split())

        for excluded in EXCLUDED_FT_DESCRIPTIONS:
            if normalized == " ".join(
                excluded.upper().split()
            ):
                return False

        return True

    return False
# ============================================================
# EXCEL ORDER LOADER
# ============================================================



def normalize_excel_value(value):

    if value is None:
        return ""

    return str(value).strip()


def load_excel_orders(excel_path):

    print()
    print("=" * 70)
    print("LOADING EXCEL ORDERS")
    print("=" * 70)

    excel_path = Path(excel_path)

    if not excel_path.exists():

        raise FileNotFoundError(
            f"Excel file was not found:\n{excel_path}"
        )

    print(
        f"Excel file: {excel_path}"
    )

    workbook = load_workbook(
        filename=excel_path,
        read_only=True,
        data_only=True
    )

    worksheet = workbook["Posted Sales Invoice Lines"]

    rows = worksheet.iter_rows(
        values_only=True
    )

    try:
        header_row = next(rows)

    except StopIteration:

        raise Exception(
            "The Excel workbook is empty."
        )

    headers = [
        normalize_excel_value(value)
        for value in header_row
    ]

    required_headers = [
    "Document No.",
    "Sell-to Customer No.",
    "Type",
    "No.",
    "Description",
    "Quantity",
    "Unit Price Incl. VAT",
    "Amount",
    "Unit of Measure Code",
    "Location Code",
]

    missing_headers = [
        header
        for header in required_headers
        if header not in headers
    ]

    if missing_headers:

        raise Exception(
            "The Excel file is missing required columns:\n"
            + "\n".join(
                f"- {header}"
                for header in missing_headers
            )
        )

    column_index = {
        header: index
        for index, header in enumerate(headers)
    }

    orders = {}

    source_rows = 0
    ignored_rows = 0
    skipped_rows = 0

    for excel_row in rows:

        if not any(
            value not in (None, "")
            for value in excel_row
        ):
            continue

        source_rows += 1

        row = {
            header: (
                excel_row[index]
                if index < len(excel_row)
                else None
            )
            for header, index in column_index.items()
        }

        # ------------------------------------------------------------
        # LOCATION CODE FILTER — FIRST GATE
        # ------------------------------------------------------------

        location_code = normalize_excel_value(
            row["Location Code"]
        )

        if not location_code.upper().startswith("K-"):

            ignored_rows += 1
            continue


        # ------------------------------------------------------------
        # ITEM TYPE FILTER
        # ------------------------------------------------------------

        row_type = normalize_excel_value(
            row["Type"]
        )

        if row_type != "Item":

            skipped_rows += 1
            continue


        # ------------------------------------------------------------
        # APPROVED DESCRIPTION FILTER
        # ------------------------------------------------------------

        description = normalize_excel_value(
            row["Description"]
        )

        if not is_approved_excel_description(
            description
        ):

            ignored_rows += 1
            continue

        document_no = normalize_excel_value(
            row["Document No."]
        )

        customer_no = normalize_excel_value(
            row["Sell-to Customer No."]
        )

        item_no = normalize_excel_value(
            row["No."]
        )

        if item_no.upper() in IGNORED_ITEM_CODES:
            ignored_rows += 1
            print(
                f"Ignoring monitored item code: {item_no}"
            )
            continue

        quantity = row["Quantity"]

        unit_price = row["Unit Price Incl. VAT"]

        amount = row["Amount"]

        unit_of_measure = normalize_excel_value(
            row["Unit of Measure Code"]
        )

        if not document_no:
            raise Exception(
                f"Blank Document No. found in Excel row "
                f"{source_rows + 1}."
            )

        if not customer_no:
            raise Exception(
                f"Blank Sell-to Customer No. found "
                f"for Document No. {document_no}."
            )

        if not item_no:
            raise Exception(
                f"Blank item number found for "
                f"Document No. {document_no}."
            )

        if not description:
            raise Exception(
                f"Blank description found for "
                f"Document No. {document_no}, "
                f"item {item_no}."
            )

        if quantity in (None, ""):

            raise Exception(
                f"Blank quantity found for "
                f"Document No. {document_no}, "
                f"item {item_no}."
            )

        if unit_price in (None, ""):

            raise Exception(
                f"Blank unit price found for "
                f"Document No. {document_no}, "
                f"item {item_no}."
            )

        try:

            quantity_value = float(quantity)

        except Exception as e:

            raise Exception(
                f"Invalid quantity '{quantity}' "
                f"for Document No. {document_no}, "
                f"item {item_no}."
            ) from e

        if quantity_value.is_integer():
            quantity_value = int(quantity_value)

        

        try:

            unit_price_value = float(
                unit_price
            )

        except Exception as e:

            raise Exception(
                f"Invalid unit price '{unit_price}' "
                f"for Document No. {document_no}, "
                f"item {item_no}."
            ) from e

        # Read the Excel Amount value explicitly.
        amount = row["Amount"]

        if amount in (None, ""):

            amount_value = None

        else:

            try:

                amount_value = float(
                    amount
                )

            except Exception as e:

                raise Exception(
                    f"Invalid Amount '{amount}' "
                    f"for Document No. {document_no}, "
                    f"item {item_no}."
                ) from e

        if document_no not in orders:

            orders[document_no] = {
                "document_no": document_no,
                "customer_no": customer_no,
                "lines": [],
            }

        existing_customer = orders[
            document_no
        ]["customer_no"]

        if existing_customer != customer_no:

            raise Exception(
                f"Document No. {document_no} contains "
                f"multiple customer numbers:\n"
                f"{existing_customer}\n"
                f"{customer_no}"
            )

        orders[
    document_no
]["lines"].append({

    "document_no": document_no,

    "customer_no": customer_no,

    "item_no": item_no,

    "description": description,

    "quantity": quantity_value,

    "unit_price_excel": unit_price_value,

    "amount_excel": amount_value,

    "unit_of_measure": unit_of_measure,

    "location_code": normalize_excel_value(
        row["Location Code"]
    ),

})

    workbook.close()

    print()
    print(
        f"Source rows read: {source_rows}"
    )

    print(
        f"FT/WT tile rows ignored: {ignored_rows}"
    )

    print(
        f"Non-item rows skipped: {skipped_rows}"
    )

    print(
        f"Usable order rows: "
        f"{sum(len(order['lines']) for order in orders.values())}"
    )

    print(
        f"Unique Document Nos.: {len(orders)}"
    )

    print()
    print(
        "EXCEL ORDER LOAD COMPLETED."
    )

    return orders
def audit_excel_pricing_codes(excel_path):

    print()
    print("=" * 80)
    print("EXCEL PRICING CODE AUDIT")
    print("=" * 80)

    workbook = load_workbook(
        filename=excel_path,
        read_only=True,
        data_only=True
    )

    worksheet = workbook.active

    rows = worksheet.iter_rows(
        values_only=True
    )

    header_row = next(rows)

    headers = [
        normalize_excel_value(value)
        for value in header_row
    ]

    column_index = {
        header: index
        for index, header in enumerate(headers)
    }

    pricing_patterns = [
        "24", "33", "44", "66",
        "36", "40", "45", "55", "61"
    ]

    results = {
        prefix: []
        for prefix in pricing_patterns
    }

    for excel_row in rows:

        if not any(
            value not in (None, "")
            for value in excel_row
        ):
            continue

        location_code = normalize_excel_value(
            excel_row[
                column_index["Location Code"]
            ]
        )

        if not location_code.upper().startswith("K-"):
            continue

        description = normalize_excel_value(
            excel_row[
                column_index["Description"]
            ]
        )

        item_no = normalize_excel_value(
            excel_row[
                column_index["No."]
            ]
        )

        # Pricing-code rules apply only to approved FT/WT tiles.
        upper_description = description.upper()

        if not (
            upper_description.startswith("FT ")
            or upper_description.startswith("WT ")
        ):
            continue

        # Ignore FT products beginning with GG or GS.
        if (
            upper_description.startswith("FT GG")
            or upper_description.startswith("FT GS")
        ):
            continue

        matches = re.findall(
            r"(?<!\d)\d{5,6}(?!\d)",
            description
        )

        for code in matches:

            for prefix in pricing_patterns:

                if code.startswith(prefix):

                    results[prefix].append(
                        {
                            "item_no": item_no,
                            "description": description,
                            "code": code,
                        }
                    )

                    break

    workbook.close()

    for prefix in pricing_patterns:

        entries = results[prefix]

        print()
        print(
            f"PREFIX {prefix} "
            f"→ RULE PRICE "
            f"{PRICE_PREFIX_RULES[prefix]}"
        )

        print(
            f"Occurrences found: {len(entries)}"
        )

        for entry in entries[:10]:

            print(
                f"  {entry['item_no']} | "
                f"{entry['code']} | "
                f"{entry['description']}"
            )

        if len(entries) > 10:

            print(
                f"  ... and "
                f"{len(entries) - 10} more."
            )

    print()
    print("=" * 80)
    print("PRICING CODE AUDIT COMPLETED")
    print("=" * 80)
def preview_excel_orders(excel_path):

    orders = load_excel_orders(
        excel_path
    )

    print()
    print("=" * 70)
    print("EXCEL ORDER PREVIEW")
    print("=" * 70)

    preview_count = min(
        len(orders),
        10
    )

    order_items = list(
        orders.values()
)

    for index in range(
        preview_count
    ):

        order = order_items[index]

        print()
        print(
            f"ORDER {index + 1}:"
        )

        print(
            f"Document No.: "
            f"{order['document_no']}"
        )

        print(
            f"Customer No.: "
            f"{order['customer_no']}"
        )

        print(
            f"Lines: "
            f"{len(order['lines'])}"
        )

        for line in order["lines"]:

            print(
                f"  {line['item_no']} | "
                f"{line['description']} | "
                f"Qty={line['quantity']} | "
                f"Excel Price={line['unit_price_excel']}"
            )

    if len(orders) > preview_count:

        print()
        print(
            f"... and "
            f"{len(orders) - preview_count} "
            f"more orders."
        )

    print()
    print(
        "EXCEL PREVIEW COMPLETED."
    )

# ============================================================
# GENERAL HELPERS
# ============================================================

def wait_for_enter(message):
    print()
    print("=" * 70)
    print(message)
    print("=" * 70)

    if os.environ.get("CRM_UPLOAD_MODE") == "1":
        print(
            "CRM upload mode detected — "
            "ending automation session without waiting for ENTER."
        )
        return

    input(
        "Press ENTER when you are finished inspecting the CRM page..."
    )

def get_next_customer_name():
    try:
        index = int(
            CUSTOMER_NAME_STATE_FILE.read_text(
                encoding="utf-8"
            ).strip()
        )
    except (FileNotFoundError, ValueError):
        index = 0

    customer_name = CUSTOMER_NAME_SEQUENCE[
        index % len(CUSTOMER_NAME_SEQUENCE)
    ]

    CUSTOMER_NAME_STATE_FILE.write_text(
        str(index + 1),
        encoding="utf-8"
    )

    return customer_name
def record_inventory_report(
    inventory_lines,
    required_quantity
):
    total_inventory = sum(
        float(line.get("inventory", 0))
        for line in inventory_lines
    )

    required_quantity = float(
        required_quantity
    )

    overall_shortage = (
        total_inventory < required_quantity
    )

    for line in inventory_lines:

        inventory = float(
            line.get("inventory", 0)
        )

        posted_quantity = float(
            line.get("allocated_quantity", 0)
        )

        # Do not report positive-inventory lines that
        # were merely excess CRM search results and deleted.
        if (
            inventory > 0
            and posted_quantity <= 0
        ):
            continue

        if inventory <= 0:
            status = "ZERO INVENTORY"

        elif overall_shortage:
            status = "INSUFFICIENT INVENTORY"

        else:
            status = "ENOUGH INVENTORY"

        POSTING_REPORT_ROWS.append(
            {
                "Customer": CURRENT_REPORT_CONTEXT[
                    "customer_name"
                ],
                "Document No.": CURRENT_REPORT_CONTEXT[
                    "document_no"
                ],
                "Excel Item No.": CURRENT_REPORT_CONTEXT[
                    "excel_item_no"
                ],
                "Excel Description": CURRENT_REPORT_CONTEXT[
                    "excel_description"
                ],
                "Excel Qty": required_quantity,
                "CRM Item No.": line.get(
                    "item_no",
                    ""
                ),
                "CRM Product": line.get(
                    "product_title",
                    ""
                ),
                "CRM Inventory": inventory,
                "Posted Qty": posted_quantity,
                "Inventory Status": status,
            }
        )
def write_posting_report():

    if not POSTING_REPORT_ROWS:
        print()
        print(
            "No inventory results were collected. "
            "Posting report was not created."
        )
        return None

    report_time = time.strftime(
        "%Y-%m-%d_%H%M%S"
    )

    report_path = (
        Path.home()
        / "Downloads"
        / f"CRM Posting Report - {report_time}.xlsx"
    )

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "Posting Report"

    headers = [
        "Customer",
        "Document No.",
        "Excel Item No.",
        "Excel Description",
        "Excel Qty",
        "CRM Item No.",
        "CRM Product",
        "CRM Inventory",
        "Posted Qty",
        "Inventory Status",
    ]

    yellow_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )
    bright_yellow_fill = PatternFill(
    fill_type="solid",
    fgColor="FFFF00"
)

    monitor_fill = PatternFill(
        fill_type="solid",
        fgColor="9DC3E6"
    )
    red_font = Font(
        color="C00000"
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    thin_border = Border(
        left=Side(style="thin", color="D9E1F2"),
        right=Side(style="thin", color="D9E1F2"),
        top=Side(style="thin", color="D9E1F2"),
        bottom=Side(style="thin", color="D9E1F2"),
    )

    worksheet.append(headers)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = thin_border

    for row_data in POSTING_REPORT_ROWS:

        worksheet.append(
            [
                row_data[header]
                for header in headers
            ]
        )

        row_number = worksheet.max_row

        status = row_data[
            "Inventory Status"
        ]

        # Zero or insufficient inventory:
        # product description is red.
        if status in {
            "ZERO INVENTORY",
            "INSUFFICIENT INVENTORY",
        }:
            worksheet.cell(
                row=row_number,
                column=4
            ).font = red_font

        # Enough or insufficient inventory:
        # quantity/inventory/status cells are yellow.
        if status in {
            "ENOUGH INVENTORY",
            "INSUFFICIENT INVENTORY",
        }:
            for column in (
                5, 8, 9, 10
            ):
                worksheet.cell(
                    row=row_number,
                    column=column
                ).fill = yellow_fill

        # Enough inventory:
        # Excel Description cell is bright yellow.
        if status == "ENOUGH INVENTORY":
            worksheet.cell(
                row=row_number,
                column=4
            ).fill = bright_yellow_fill

    for row in worksheet.iter_rows():

        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    worksheet.freeze_panes = "A2"

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    widths = {
        "A": 16,
        "B": 16,
        "C": 16,
        "D": 42,
        "E": 12,
        "F": 16,
        "G": 44,
        "H": 16,
        "I": 14,
        "J": 24,
    }

    for column, width in widths.items():
        worksheet.column_dimensions[
            column
        ].width = width
        # ----------------------------------------------------
        # ITEMS UNDER MONITORING
        # ----------------------------------------------------
        monitor_start_row = worksheet.max_row + 2

        worksheet.merge_cells(
            start_row=monitor_start_row,
            start_column=1,
            end_row=monitor_start_row,
            end_column=4
        )

        monitor_title = worksheet.cell(
            row=monitor_start_row,
            column=1
        )

        monitor_title.value = "Items Under Monitoring"
        monitor_title.font = Font(
            bold=True,
            color="FFFFFF"
        )
        monitor_title.fill = header_fill
        monitor_title.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        monitor_headers = [
            "Item Code",
            "Status",
            "Action",
            "Note",
        ]

        for column, header in enumerate(
            monitor_headers,
            start=1
        ):
            cell = worksheet.cell(
                row=monitor_start_row + 1,
                column=column
            )
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = thin_border

        monitor_notes = (
            "Excluded from CRM posting run"
        )

        for offset, item_code in enumerate(
            sorted(IGNORED_ITEM_CODES),
            start=monitor_start_row + 2
        ):
            monitor_values = [
                item_code,
                "UNDER MONITORING",
                "NOT POSTED",
                monitor_notes,
            ]

            for column, value in enumerate(
                monitor_values,
                start=1
            ):
                cell = worksheet.cell(
                    row=offset,
                    column=column
                )
                cell.value = value
                cell.fill = monitor_fill
                cell.border = thin_border
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )
    worksheet.row_dimensions[1].height = 32

    workbook.save(
        report_path
    )

    print()
    print("=" * 70)
    print("CRM POSTING REPORT CREATED")
    print("=" * 70)
    print(
        f"Report: {report_path}"
    )

    return report_path
def visible_count(locator):
    count = locator.count()
    visible = 0

    for i in range(count):
        try:
            if locator.nth(i).is_visible():
                visible += 1
        except Exception:
            pass

    return visible


def find_order_row(page, product_title):

    # --------------------------------------------------------
    # First try the exact product title.
    # This is the normal path and preserves the behaviour
    # that is already working for the WC-001W products.
    # --------------------------------------------------------

    rows = page.locator("tr").filter(
        has_text=product_title
    )

    count = rows.count()

    if count > 0:
        for i in range(count):
            row = rows.nth(i)

            try:
                if row.is_visible():
                    return row
            except Exception:
                pass

    # --------------------------------------------------------
    # Fallback:
    # Some CRM rows split the product code and description
    # into separate text blocks.
    #
    # Example:
    #
    # 311002539
    #
    # PMCP24101J WH WALLTILE 2540 15PCS J1
    #
    # In that situation, the complete product_title may not
    # match as one string.
    # --------------------------------------------------------

    print()
    print(
        "Exact order-line text was not matched."
    )
    print(
        "Inspecting visible order rows for a partial match..."
    )

    visible_rows = page.locator("tr:visible")

    visible_count = visible_rows.count()

    print(
        f"Visible order rows found: {visible_count}"
    )

    # Extract useful identifying parts from the title.
    #
    # For the 24101 product, the first part is:
    # 311002539
    #
    # For other products, the full title may still work
    # through the exact-match path above.

    title_parts = product_title.split()

    possible_identifiers = []

    if len(title_parts) > 0:
        possible_identifiers.append(
            title_parts[0]
        )

    # Also use the product code embedded in the title
    # when the title contains a recognizable numeric code.
    for part in title_parts:
        if part.isdigit() and len(part) >= 5:
            if part not in possible_identifiers:
                possible_identifiers.append(part)

    for i in range(visible_count):

        row = visible_rows.nth(i)

        try:
            row_text = row.inner_text().strip()

            for identifier in possible_identifiers:

                if identifier in row_text:

                    print()
                    print(
                        "POSSIBLE ORDER LINE MATCH:"
                    )
                    print(ascii(row_text))

                    return row

        except Exception:
            continue

    raise Exception(
        f"Order line not found:\n"
        f"{product_title}"
    )

def enter_quantity(page, product_title, quantity):

    print()
    print("=" * 60)
    print("ENTERING QUANTITY")
    print("=" * 60)
    print(f"Product: {product_title}")
    print(f"Quantity: {quantity}")

    # Find the correct CRM order line.
    row = find_order_row(
        page,
        product_title
    )

    # The CRM quantity field is an input[type="number"]
    # inside the product's order row.
    quantity_input = row.locator(
        'input[type="number"]'
    ).first

    if quantity_input.count() == 0:
        raise Exception(
            f"Quantity input not found for:\n"
            f"{product_title}"
        )

    quantity_input.scroll_into_view_if_needed()

    # Clear any existing value and enter the quantity.
    quantity_input.fill(
        str(quantity)
    )

    page.wait_for_timeout(300)

    entered_quantity = quantity_input.input_value()

    print(
        f"Quantity field now shows: "
        f"'{entered_quantity}'"
    )

    if entered_quantity != str(quantity):
        raise Exception(
            f"Quantity was not entered correctly for:\n"
            f"{product_title}\n"
            f"Expected: {quantity}\n"
            f"Found: {entered_quantity}"
        )

    print(
        "QUANTITY ENTERED SUCCESSFULLY"
    )
def enter_unit_price(page, product_title, unit_price):

    print()
    print("=" * 60)
    print("ENTERING UNIT PRICE")
    print("=" * 60)
    print(f"Product: {product_title}")
    print(f"Unit price: {unit_price}")

    # Find the correct CRM order line.
    row = find_order_row(
        page,
        product_title
    )

    # The CRM unit-price field is an input[type="number"]
    # with the placeholder "Please input unit price".
    price_input = row.locator(
        'input[placeholder="Please input unit price"]'
    ).first

    if price_input.count() == 0:
        raise Exception(
            f"Unit price input not found for:\n"
            f"{product_title}"
        )

    price_input.scroll_into_view_if_needed()

    # Clear any existing value and enter the unit price.
    price_input.fill(
        str(unit_price)
    )

    page.wait_for_timeout(300)

    entered_price = price_input.input_value()

    print(
        f"Unit price field now shows: "
        f"'{entered_price}'"
    )

    if entered_price != str(unit_price):
        raise Exception(
            f"Unit price was not entered correctly for:\n"
            f"{product_title}\n"
            f"Expected: {unit_price}\n"
            f"Found: {entered_price}"
        )

    print(
        "UNIT PRICE ENTERED SUCCESSFULLY"
    ) 
      
# ============================================================
# WAREHOUSE / INVENTORY
# ============================================================

def debug_visible_warehouse_dropdown(page):
    """
    Prints the actual visible dropdown structure so we can see
    exactly how Element Plus is rendering the warehouse options.
    """

    print()
    print("========== WAREHOUSE DROPDOWN DEBUG ==========")

    try:
        dropdowns = page.locator(
            ".el-select-dropdown:visible"
        )

        print(
            f"Visible Element Plus dropdowns: {dropdowns.count()}"
        )

        for i in range(dropdowns.count()):
            dropdown = dropdowns.nth(i)

            print()
            print(f"--- Dropdown {i} ---")

            try:
                print("TEXT:")
                print(dropdown.inner_text())
            except Exception as e:
                print(
                    f"Could not read dropdown text: {e}"
                )

            try:
                print()
                print("HTML:")

                html = dropdown.evaluate(
                    "(el) => el.outerHTML"
                )

                print(html[:12000])

            except Exception as e:
                print(
                    f"Could not read dropdown HTML: {e}"
                )

    except Exception as e:
        print(
            f"Could not inspect Element Plus dropdowns: {e}"
        )

    try:
        suam_matches = page.get_by_text(
            WAREHOUSE_NAME,
            exact=False
        )

        print()
        print(
            f"Visible text matches containing "
            f"'{WAREHOUSE_NAME}': "
            f"{visible_count(suam_matches)}"
        )

        for i in range(suam_matches.count()):
            candidate = suam_matches.nth(i)

            try:
                if candidate.is_visible():

                    print()
                    print(
                        f"--- Visible SUAM match {i} ---"
                    )

                    print(
                        "Tag:",
                        candidate.evaluate(
                            "(el) => el.tagName"
                        )
                    )

                    print(
                        "Text:",
                        candidate.inner_text()
                    )

                    print(
                        "Class:",
                        candidate.get_attribute("class")
                    )

                    print(
                        "Outer HTML:",
                        candidate.evaluate(
                            "(el) => el.outerHTML"
                        )[:5000]
                    )

            except Exception:
                pass

    except Exception as e:
        print(
            f"Additional SUAM STORES inspection failed: {e}"
        )

    print()
    print("========== END WAREHOUSE DEBUG ==========")


def select_suam_stores(page, warehouse_select):
    """
    Select SUAM STORES from the warehouse dropdown.

    read_inventory() has already opened the dropdown before
    calling this function, so this function must NOT click
    the warehouse selector again.
    """

    print()
    print("=" * 60)
    print("SELECTING WAREHOUSE")
    print("=" * 60)

    print(f"Target warehouse: {WAREHOUSE_NAME}")
    print("Looking for the visible SUAM STORES option...")

    page.wait_for_timeout(800)

    # The CRM uses Element Plus and creates multiple dropdown
    # instances in the DOM. Some are hidden.
    #
    # IMPORTANT:
    # :visible ensures we only target the currently open
    # dropdown option.
    visible_options = page.locator(
        "li.el-select-dropdown__item:visible"
    ).filter(
        has_text=WAREHOUSE_NAME
    )

    count = visible_options.count()

    print(
        f"Visible SUAM STORES options found: {count}"
    )

    if count == 0:
        print()
        print("No visible SUAM STORES option found.")

        print()
        print("Currently visible dropdown options:")

        visible_dropdown_options = page.locator(
            "li.el-select-dropdown__item:visible"
        )

        visible_count = visible_dropdown_options.count()

        for i in range(visible_count):
            try:
                text = (
                    visible_dropdown_options
                    .nth(i)
                    .inner_text()
                    .strip()
                )

                print(
                    f"Visible option {i}: '{text}'"
                )

            except Exception:
                pass

        raise Exception(
            "Visible SUAM STORES dropdown option was not found."
        )

    # Use the FIRST visible SUAM STORES option.
    option = visible_options.first

    try:
        print()
        print("Actual visible SUAM STORES option found.")

        print(
            "Option HTML:"
        )

        print(
            option.evaluate(
                "(el) => el.outerHTML"
            )[:3000]
        )

    except Exception:
        pass

    # Make sure the option is actually visible.
    option.wait_for(
        state="visible",
        timeout=5000
    )

    # Scroll the actual option into view.
    option.scroll_into_view_if_needed()

    page.wait_for_timeout(300)

    # --------------------------------------------------------
    # NORMAL CLICK
    # --------------------------------------------------------

    try:
        print()
        print("Clicking visible SUAM STORES option...")

        option.click(
            timeout=5000
        )

        print(
            "SUAM STORES option clicked."
        )

    except Exception as e:

        print(
            f"Normal click failed: {e}"
        )

        # ----------------------------------------------------
        # FORCE CLICK FALLBACK
        # ----------------------------------------------------

        try:
            print(
                "Trying force click..."
            )

            option.click(
                force=True,
                timeout=5000
            )

            print(
                "Force click executed."
            )

        except Exception as e2:

            print(
                f"Force click failed: {e2}"
            )

            # ------------------------------------------------
            # DOM CLICK FALLBACK
            # ------------------------------------------------

            try:
                print(
                    "Trying direct DOM click..."
                )

                option.evaluate(
                    "(el) => el.click()"
                )

                print(
                    "DOM click executed."
                )

            except Exception as e3:

                raise Exception(
                    "All SUAM STORES click attempts failed."
                ) from e3

    # Give Vue/Element Plus time to update the field.
    page.wait_for_timeout(1000)

    # --------------------------------------------------------
    # VERIFY SELECTION
    # --------------------------------------------------------

    try:
        current_value = (
            warehouse_select
            .inner_text()
            .strip()
        )

    except Exception:
        current_value = ""

    print()
    print(
        f"Warehouse field now shows: '{current_value}'"
    )

    if (
        WAREHOUSE_NAME.upper()
        in current_value.upper()
    ):

        print()
        print("=" * 60)
        print("WAREHOUSE SELECTED SUCCESSFULLY")
        print("=" * 60)

        print(
            f"Confirmed warehouse: {WAREHOUSE_NAME}"
        )

        return True

    # Secondary verification directly from the select.
    try:

        selected_text = (
            warehouse_select
            .locator(
                ".el-select__selected-item, "
                ".el-select__placeholder"
            )
            .last
            .inner_text()
            .strip()
        )

        print(
            f"Selected-item text: '{selected_text}'"
        )

        if (
            WAREHOUSE_NAME.upper()
            in selected_text.upper()
        ):

            print()
            print("=" * 60)
            print("WAREHOUSE SELECTED SUCCESSFULLY")
            print("=" * 60)

            print(
                f"Confirmed warehouse: {WAREHOUSE_NAME}"
            )

            return True

    except Exception as e:
        print(
            f"Secondary verification failed: {e}"
        )

    raise Exception(
        f"Could not select {WAREHOUSE_NAME}."
    )



def read_inventory(page, product_title):

    print()
    print("Processing warehouse for:")
    print(product_title)

    # --------------------------------------------------------
    # Find the product row.
    # The CRM may re-render the table, so always locate it fresh.
    # --------------------------------------------------------
    row = find_order_row(
        page,
        product_title
    )

    cells = row.locator("td")
    cell_count = cells.count()

    print(
        f"Number of cells in row: {cell_count}"
    )

    if cell_count < 5:
        raise Exception(
            f"Unexpected order row structure for "
            f"{product_title}. "
            f"Only {cell_count} cells found."
        )

    # Confirmed CRM table structure:
    #
    # 0 = Item number
    # 1 = Product title
    # 2 = Unit
    # 3 = Warehouse
    # 4 = Dealer warehouse inventory

    warehouse_cell = cells.nth(3)

    warehouse_select = warehouse_cell.locator(
        ".el-select"
    ).first

    if warehouse_select.count() == 0:
        raise Exception(
            f"Warehouse selector not found for:\n"
            f"{product_title}"
        )

    try:
        current_warehouse = (
            warehouse_select.inner_text().strip()
        )
    except Exception:
        current_warehouse = ""

    print(
        f"Current warehouse: {current_warehouse}"
    )

    if WAREHOUSE_NAME.upper() not in current_warehouse.upper():

        print()
        print("OPENING WAREHOUSE DROPDOWN...")

        warehouse_select.click(force=True)

        selected = select_suam_stores(
            page,
            warehouse_select
        )

        if not selected:
            raise Exception(
                f"Could not select {WAREHOUSE_NAME}."
            )

        print()
        print("Warehouse selection completed.")

    else:
        print(
            "Warehouse is already SUAM STORES."
        )

    # --------------------------------------------------------
    # IMPORTANT:
    # Selecting the warehouse causes the CRM/Vue table
    # to re-render. The old row locator can therefore become
    # invalid or temporarily disappear.
    #
    # Wait until the product row is available again, then
    # locate the fresh row.
    # --------------------------------------------------------

    print()
    print(
        "Waiting for CRM table to finish re-rendering..."
    )

    page.wait_for_timeout(1500)

    row = page.locator("tr").filter(
        has_text=product_title
    ).first

    try:
        row.wait_for(
            state="visible",
            timeout=10000
        )
    except Exception:
        raise Exception(
            f"Product row did not reappear after "
            f"warehouse selection:\n{product_title}"
        )

    print(
        "Product row is visible again."
    )

    # --------------------------------------------------------
    # Re-read the cells from the fresh row.
    # --------------------------------------------------------

    cells = row.locator("td")

    cell_count = cells.count()

    print(
        f"Refreshed row cell count: {cell_count}"
    )

    if cell_count < 5:
        raise Exception(
            f"Unexpected refreshed row structure for "
            f"{product_title}. "
            f"Only {cell_count} cells found."
        )

    inventory_cell = cells.nth(4)

    # Wait briefly for the inventory value to populate.
    page.wait_for_timeout(500)

    inventory_text = (
        inventory_cell.inner_text().strip()
    )

    print(
        f"Dealer warehouse inventory: "
        f"'{inventory_text}'"
    )

    return inventory_text


# ============================================================
# PRODUCT WINDOW
# ============================================================

def open_product_window(page):

    print()
    print("=" * 50)
    print("OPENING ADD PRODUCT WINDOW")
    print("=" * 50)

    selectors = [
        'button:has-text("Add a product")',
        'button:has-text("Add Product")',
        '.el-button:has-text("Add a product")',
        '.el-button:has-text("Add Product")',
    ]

    for selector in selectors:

        try:

            buttons = page.locator(selector)
            count = buttons.count()

            print(
                f"Selector '{selector}' found "
                f"{count} element(s)."
            )

            for i in range(count):

                button = buttons.nth(i)

                try:

                    if not button.is_visible():
                        continue

                    print(
                        f"Visible Add Product button found "
                        f"using: {selector}"
                    )

                    try:
                        print(
                            f"Button text: "
                            f"'{button.inner_text().strip()}'"
                        )
                    except Exception:
                        pass

                    button.scroll_into_view_if_needed()

                    page.wait_for_timeout(300)

                    button.click(force=True)

                    page.wait_for_timeout(1000)

                    print(
                        "Add a product button clicked successfully."
                    )

                    return

                except Exception as e:

                    print(
                        f"Could not click candidate {i}: {e}"
                    )

        except Exception as e:

            print(
                f"Selector failed: {selector}"
            )

            print(e)

    # --------------------------------------------------------
    # Diagnostic output
    # --------------------------------------------------------

    print()
    print("=" * 50)
    print("ADD PRODUCT BUTTON NOT FOUND")
    print("=" * 50)

    print()
    print(
        "Visible buttons currently on the order page:"
    )

    try:

        visible_buttons = page.locator(
            "button:visible"
        )

        count = visible_buttons.count()

        print(
            f"Visible buttons found: {count}"
        )

        for i in range(count):

            try:

                button = visible_buttons.nth(i)

                text = button.inner_text().strip()

                if text:

                    print(
                        f"BUTTON {i}: '{text}'"
                    )

            except Exception:
                continue

    except Exception as e:

        print(
            f"Could not inspect visible buttons: {e}"
        )

    print()
    print(
        "The CRM page will remain open for inspection."
    )

    raise Exception(
        "Could not locate the Add a product button."
    )


def search_product(page, search_text):

    print(
        f"Searching product: {search_text}"
    )

    print(
        "Finding 'The title of the product' search field..."
    )

    search_input = page.locator(
        'input[placeholder="Please input The title of the product"]'
    ).first

    search_input.wait_for(
        state="visible",
        timeout=15000
    )

    print(
        "Correct title search field found."
    )

    search_input.fill(search_text)

    search_button = page.get_by_role(
        "button",
        name="Search",
        exact=True
    )

    print(
        "Clicking Search..."
    )

    search_button.click(force=True)

    page.wait_for_timeout(3000)

    print(
        f"Search completed for: {search_text}"
    )


# ============================================================
# PRODUCT SELECTION
# ============================================================

def extract_embedded_crm_code(text):
    """
    Extract a CRM-style product code such as:

        KF-028
        KS-7843S
        KS-1050
        WC-793P
        WC-009
        WT-02D

    Returns the first matching code, or an empty string.
    """

    text = str(text).upper()

    match = re.search(
    r"\b[A-Z]{2}-\d{2,5}[A-Z]{0,2}\b",
    text
)

    if match:
        return match.group(0)

    return ""
def select_all_product_results(page):

    print()
    print("=" * 60)
    print("SELECTING ALL CRM PRODUCT RESULTS")
    print("=" * 60)

    rows = page.locator(
        "tr:visible"
    )

    row_count = rows.count()

    print(
        f"Visible product rows found: {row_count}"
    )

    selected_products = []

    for i in range(row_count):

        row = rows.nth(i)

        try:

            cells = row.locator("td")

            # Header rows do not have product checkbox
            # inputs inside table cells.
            checkbox_input = row.locator(
            'input[type="checkbox"]'
            ).first

            visible_checkbox = row.locator(
                ".el-checkbox__inner"
            ).first

            if visible_checkbox.count() == 0:
                continue

            # ------------------------------------------------
            # READ PRODUCT IDENTITY
            # ------------------------------------------------

            item_no = ""

            product_title = ""

            if cells.count() >= 1:

                item_no = " ".join(
                    cells.nth(0)
                    .inner_text()
                    .split()
                ).strip()

            if cells.count() >= 2:

                product_title = " ".join(
                    cells.nth(1)
                    .inner_text()
                    .split()
                ).strip()

            row_text = " ".join(
                row.inner_text().split()
            ).strip()

            # Skip the CRM table header row.
            if (
                "THE TITLE OF THE PRODUCT" in row_text.upper()
                and "STANDARD PACKING UNITS" in row_text.upper()
            ):
                continue

            if not product_title:

                product_title = row_text

            print()
            print(
                f"PRODUCT RESULT {len(selected_products) + 1}:"
            )

            print(
                f"  Item No.: {item_no}"
            )

            print(
                f"  Product: {product_title}"
            )

            print(
                f"  Row: {ascii(row_text)}"
            )

            # ------------------------------------------------
            # CHECK CURRENT STATE
            # ------------------------------------------------

            try:

                if checkbox_input.count() > 0:
                    already_checked = checkbox_input.is_checked()
                else:
                    already_checked = (
                        "is-checked"
                        in (
                            visible_checkbox
                            .locator("..")
                            .get_attribute("class")
                            or ""
                        )
                    )

            except Exception:

             already_checked = False

            print(
                f"  Already selected: "
                f"{already_checked}"
            )

            if not already_checked:

                visible_checkbox.scroll_into_view_if_needed()

                visible_checkbox.click(
                    force=True
                )

                page.wait_for_timeout(200)

            try:

                if checkbox_input.count() > 0:
                    final_state = checkbox_input.is_checked()
                else:
                    final_state = (
                        "is-checked"
                        in (
                            visible_checkbox
                            .locator("..")
                            .get_attribute("class")
                            or ""
                        )
                    )

            except Exception:

                final_state = False

            if not final_state:

                raise Exception(
                    f"Product checkbox could not be "
                    f"selected:\n{row_text}"
                )

            selected_products.append(
                {
                    "item_no": item_no,
                    "product_title": product_title,
                }
            )

        except Exception as e:

            raise Exception(
                f"Could not process product result "
                f"row {i}: {e}"
            ) from e

    if not selected_products:

        raise Exception(
            "No selectable CRM product results "
            "were found after the search."
        )

    print()
    print(
        f"Total CRM product results selected: "
        f"{len(selected_products)}"
    )

    print()
    print(
        "ALL CRM PRODUCT RESULTS SELECTED."
    )

    return selected_products


def confirm_product_selection(page):

    print()
    print(
        "Confirming product selection..."
    )

    confirm_buttons = page.get_by_role(
        "button",
        name="Confirm",
        exact=True
    )

    count = confirm_buttons.count()

    visible_confirm = None

    for i in range(count):

        button = confirm_buttons.nth(i)

        try:

            if button.is_visible():

                visible_confirm = button
                break

        except Exception:
            pass

    if visible_confirm is None:

        raise Exception(
            "Visible Confirm button was not found "
            "after product selection."
        )

    visible_confirm.scroll_into_view_if_needed()

    visible_confirm.click(force=True)

    page.wait_for_timeout(1500)

    print(
        "Product selection confirmed."
    )


# ============================================================
# MENU / DELETE
# ============================================================

def inspect_menu(page, product_title):

    print()
    print("Inspecting Menu for:")
    print(product_title)

    row = find_order_row(
        page,
        product_title
    )

    cells = row.locator("td")

    cell_count = cells.count()

    print(
        f"Number of cells: {cell_count}"
    )

    if cell_count == 0:

        raise Exception(
            f"No cells found for {product_title}"
        )

    # Menu is the last column
    menu_cell = cells.nth(
        cell_count - 1
    )

    try:
        menu_text = menu_cell.inner_text().strip()
    except Exception:
        menu_text = ""

    print(
        f"Menu cell text: '{menu_text}'"
    )

    delete_text = menu_cell.get_by_text(
        "Delete",
        exact=True
    )

    delete_count = 0

    for i in range(delete_text.count()):

        try:

            if delete_text.nth(i).is_visible():

                delete_count += 1

        except Exception:
            pass

    print(
        f"Delete elements found: {delete_count}"
    )

    if delete_count > 0:

        print(
            f"DELETE AVAILABLE for: "
            f"{product_title}"
        )

    else:

        print(
            f"DELETE NOT AVAILABLE for: "
            f"{product_title}"
        )

    return {
        "row": row,
        "menu_cell": menu_cell,
        "menu_text": menu_text,
        "delete_text": delete_text,
        "delete_count": delete_count,
    }


def delete_product_line(page, product_title):

    print()
    print("Attempting to delete:")
    print(product_title)

    info = inspect_menu(
        page,
        product_title
    )

    if info["delete_count"] == 0:

        print()
        print(
            "DELETE IS NOT AVAILABLE."
        )

        print(
            "The CRM does not currently provide a "
            "Delete option for this line."
        )

        return False

    delete_text = info["delete_text"]

    visible_delete = None

    for i in range(delete_text.count()):

        candidate = delete_text.nth(i)

        try:

            if candidate.is_visible():

                visible_delete = candidate
                break

        except Exception:
            pass

    if visible_delete is None:

        print(
            "Delete was detected but no visible Delete "
            "element could be clicked."
        )

        return False

    print()
    print(
        "Delete option is available."
    )

    visible_delete.scroll_into_view_if_needed()

    visible_delete.click(force=True)

    page.wait_for_timeout(500)

    print(
        "Checking for deletion confirmation..."
    )

    confirmation_found = False

    try:

        confirm_buttons = page.get_by_role(
            "button",
            name="Confirm",
            exact=True
        )

        button_count = confirm_buttons.count()

        for i in range(button_count):

            button = confirm_buttons.nth(i)

            try:

                if button.is_visible():

                    print(
                        "Deletion confirmation dialog detected."
                    )

                    button.click(force=True)

                    confirmation_found = True

                    print(
                        "Deletion confirmed."
                    )

                    break

            except Exception:
                continue

    except Exception as e:

        print(
            f"Confirmation check produced: {e}"
        )

    if not confirmation_found:

        print(
            "No deletion confirmation dialog detected."
        )

    print(
        "Waiting for product line to disappear..."
    )

    page.wait_for_timeout(1000)

    try:

        page.locator("tr").filter(
            has_text=product_title
        ).first.wait_for(
            state="detached",
            timeout=10000
        )

        print()
        print(
            f"Successfully deleted: "
            f"{product_title}"
        )

        return True

    except PlaywrightTimeoutError:

        remaining = (
            page.locator("tr")
            .filter(has_text=product_title)
            .count()
        )

        if remaining == 0:

            print()
            print(
                f"Successfully deleted: "
                f"{product_title}"
            )

            return True

        print()
        print(
            f"WARNING: Could not verify deletion of "
            f"{product_title}"
        )

        return False


# ============================================================
# ORDER LINE INSPECTION
# ============================================================

def print_current_order_lines(page):

    print()
    print("=" * 70)
    print("CURRENT CRM ORDER LINES")
    print("=" * 70)

    rows = page.locator("tr")

    found = 0

    for i in range(rows.count()):

        row = rows.nth(i)

        try:

            if not row.is_visible():
                continue

            cells = row.locator("td")

            if cells.count() < 5:
                continue

            text = row.inner_text().strip()

            if not text:
                continue

            print()
            print(
                f"ROW {i}:"
            )

            print(text)

            found += 1

        except Exception:
            continue

    print()

    print(
        f"Visible table rows inspected: {found}"
    )

def normalize_order_line_text(value):

    return " ".join(
        str(value).split()
    ).strip()


def get_order_line_records(page):

    rows = page.locator(
        "tr:visible"
    )

    records = []

    for i in range(rows.count()):

        row = rows.nth(i)

        try:

            cells = row.locator("td")

            if cells.count() < 5:
                continue

            item_no = normalize_order_line_text(
                cells.nth(0).inner_text()
            )

            product_title = normalize_order_line_text(
                cells.nth(1).inner_text()
            )

            if not item_no or not product_title:
                continue

            records.append(
                {
                    "item_no": item_no,
                    "product_title": product_title,
                    "key": (
                        item_no.upper(),
                        product_title.upper(),
                    ),
                }
            )

        except Exception:
            continue

    return records
def find_order_rows_by_identity(
    page,
    item_no,
    product_title
):

    target_item = normalize_order_line_text(
        item_no
    ).upper()

    target_title = normalize_order_line_text(
        product_title
    ).upper()

    rows = page.locator(
        "tr:visible"
    )

    matches = []

    for i in range(rows.count()):

        row = rows.nth(i)

        try:

            cells = row.locator("td")

            if cells.count() < 5:
                continue

            current_item = normalize_order_line_text(
                cells.nth(0).inner_text()
            ).upper()

            current_title = normalize_order_line_text(
                cells.nth(1).inner_text()
            ).upper()

            if (
                current_item == target_item
                and current_title == target_title
            ):

                matches.append(row)

        except Exception:
            continue

    return matches
def set_order_line_quantity(
    page,
    item_no,
    product_title,
    occurrence,
    quantity
):

    rows = find_order_rows_by_identity(
        page,
        item_no,
        product_title
    )

    if occurrence >= len(rows):

        raise Exception(
            f"Could not find CRM line occurrence "
            f"{occurrence} for:\n"
            f"{item_no} | {product_title}"
        )

    row = rows[occurrence]

    quantity_input = row.locator(
        'input[type="number"]'
    ).first

    if quantity_input.count() == 0:

        raise Exception(
            f"Quantity input not found for:\n"
            f"{item_no} | {product_title}"
        )

    quantity_input.scroll_into_view_if_needed()

    quantity_input.fill(
        str(quantity)
    )

    page.wait_for_timeout(300)

    entered = quantity_input.input_value()

    print(
        f"  CRM quantity set to: {entered}"
    )

    try:

        entered_numeric = float(entered)
        expected_numeric = float(quantity)

    except Exception:

        raise Exception(
            f"Could not verify quantity for:\n"
            f"{item_no} | {product_title}"
        )

    if entered_numeric != expected_numeric:

        raise Exception(
            f"Quantity mismatch for:\n"
            f"{item_no} | {product_title}\n"
            f"Expected: {quantity}\n"
            f"Found: {entered}"
        )
def set_order_line_unit_price(
    page,
    item_no,
    product_title,
    occurrence,
    unit_price
):

    rows = find_order_rows_by_identity(
        page,
        item_no,
        product_title
    )

    if occurrence >= len(rows):

        raise Exception(
            f"Could not find CRM line occurrence "
            f"{occurrence} for:\n"
            f"{item_no} | {product_title}"
        )

    row = rows[occurrence]

    price_input = row.locator(
        'input[placeholder="Please input unit price"]'
    ).first

    if price_input.count() == 0:

        raise Exception(
            f"Unit price input not found for:\n"
            f"{item_no} | {product_title}"
        )

    price_input.scroll_into_view_if_needed()

    price_input.fill(
        str(unit_price)
    )

    page.wait_for_timeout(300)

    entered = price_input.input_value()

    print(
        f"  CRM unit price set to: {entered}"
    )

    try:

        entered_numeric = float(
            entered
        )

        expected_numeric = float(
            unit_price
        )

    except Exception:

        raise Exception(
            f"Could not verify unit price for:\n"
            f"{item_no} | {product_title}"
        )

    if entered_numeric != expected_numeric:

        raise Exception(
            f"Unit price mismatch for:\n"
            f"{item_no} | {product_title}\n"
            f"Expected: {unit_price}\n"
            f"Found: {entered}"
        )
def delete_order_line_occurrence(
    page,
    item_no,
    product_title,
    occurrence
):

    rows = find_order_rows_by_identity(
        page,
        item_no,
        product_title
    )

    if occurrence >= len(rows):

        print(
            f"Line occurrence {occurrence} no longer exists:"
            f" {item_no} | {product_title}"
        )

        return False

    row = rows[occurrence]

    cells = row.locator("td")

    menu_cell = cells.nth(
        cells.count() - 1
    )

    delete_button = menu_cell.get_by_text(
        "Delete",
        exact=True
    )

    visible_delete = None

    for i in range(delete_button.count()):

        candidate = delete_button.nth(i)

        try:

            if candidate.is_visible():

                visible_delete = candidate
                break

        except Exception:
            continue

    if visible_delete is None:

        print()
        print(
            "DELETE NOT AVAILABLE"
        )

        print(
            f"Line: {item_no} | {product_title}"
        )

        return False

    print()
    print(
        "Deleting CRM line:"
    )

    print(
        f"  {item_no} | {product_title}"
    )

    print(
        f"  Occurrence: {occurrence}"
    )

    visible_delete.scroll_into_view_if_needed()

    visible_delete.click(
        force=True
    )

    page.wait_for_timeout(400)

    # --------------------------------------------------------
    # CONFIRM DELETION IF REQUIRED
    # --------------------------------------------------------

    confirm_buttons = page.get_by_role(
        "button",
        name="Confirm",
        exact=True
    )

    for i in range(confirm_buttons.count()):

        button = confirm_buttons.nth(i)

        try:

            if button.is_visible():

                button.click(
                    force=True
                )

                print(
                    "Deletion confirmed."
                )

                break

        except Exception:
            continue

    page.wait_for_timeout(800)

    remaining_rows = find_order_rows_by_identity(
        page,
        item_no,
        product_title
    )

    if len(remaining_rows) < len(rows):

        print(
            "CRM line deleted successfully."
        )

        return True

    print(
        "CRM line could not be verified as deleted."
    )

    return False
def allocate_quantity_across_new_lines(
    page,
    new_lines,
    required_quantity,
    unit_price
):

    print()
    print("=" * 70)
    print("ALLOCATING BC QUANTITY ACROSS CRM LINES")
    print("=" * 70)

    required_quantity = float(
        required_quantity
    )

    if required_quantity <= 0:
        raise Exception(
            f"Invalid required quantity: "
            f"{required_quantity}"
        )

    if not new_lines:
        raise Exception(
            "No newly created CRM lines were found "
            "for the current BC item."
        )

    print()
    print(
        f"BC required quantity: {required_quantity}"
    )

    print(
        f"New CRM lines to inspect: {len(new_lines)}"
    )

    # --------------------------------------------------------
    # READ INVENTORY FROM EVERY NEW CRM LINE
    # --------------------------------------------------------

    inventory_lines = []

    for line in new_lines:

        inventory = read_inventory_for_order_line(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"]
        )

        line_copy = dict(line)

        line_copy["inventory"] = float(
            inventory
        )

        line_copy["allocated_quantity"] = 0

        inventory_lines.append(
            line_copy
        )

    # --------------------------------------------------------
    # SEPARATE ZERO-STOCK AND POSITIVE-STOCK LINES
    # --------------------------------------------------------

    zero_lines = [
        line
        for line in inventory_lines
        if line["inventory"] <= 0
    ]

    positive_lines = [
        line
        for line in inventory_lines
        if line["inventory"] > 0
    ]

    print()
    print(
        f"Zero-inventory lines: {len(zero_lines)}"
    )

    print(
        f"Positive-inventory lines: {len(positive_lines)}"
    )

    # --------------------------------------------------------
    # HIGHEST INVENTORY FIRST
    # --------------------------------------------------------

    positive_lines.sort(
        key=lambda line: line["inventory"],
        reverse=True
    )

    # --------------------------------------------------------
    # ALLOCATE ONE SHARED BC QUANTITY
    # --------------------------------------------------------

    remaining = required_quantity
    retained_lines = []
    excess_lines = []

    for line in positive_lines:

        if remaining > 0:

            quantity_to_use = min(
                line["inventory"],
                remaining
            )

            line["allocated_quantity"] = (
                quantity_to_use
            )

            retained_lines.append(
                line
            )

            remaining -= quantity_to_use

        else:

            excess_lines.append(
                line
            )

    print()
    print(
        f"Total positive inventory available: "
        f"{sum(line['inventory'] for line in positive_lines)}"
    )

    print()
    print(
        f"Remaining BC quantity after allocation: "
        f"{remaining}"
    )

    

    # --------------------------------------------------------
    # SET QUANTITY ON RETAINED LINES
    # --------------------------------------------------------

    print()
    print(
        "SETTING QUANTITIES ON RETAINED CRM LINES"
    )

    for line in retained_lines:

        print()
        print(
            f"Item: {line['item_no']}"
        )

        print(
            f"Product: {line['product_title']}"
        )

        print(
            f"Inventory: {line['inventory']}"
        )

        print(
            f"Allocated quantity: "
            f"{line['allocated_quantity']}"
        )

        set_order_line_quantity(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"],
            line["allocated_quantity"]
        )
        set_order_line_unit_price(
        page,
        line["item_no"],
        line["product_title"],
        line["occurrence"],
        unit_price
    )
    # --------------------------------------------------------
    # DELETE ZERO-INVENTORY AND UNUSED LINES
    # --------------------------------------------------------

    lines_to_delete = (
        zero_lines
        + excess_lines
    )

    # For the same product identity, delete the highest
    # occurrence first so lower occurrence numbers do not shift
    # before they are processed.
    lines_to_delete.sort(
        key=lambda line: (
            line["key"],
            line["occurrence"]
        ),
        reverse=True
    )

    print()
    print(
        "REMOVING ZERO-INVENTORY / EXCESS CRM LINES"
    )

    for line in lines_to_delete:

        delete_order_line_occurrence(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"]
        )

    # --------------------------------------------------------
    # SHORTFALL REPORT
    # --------------------------------------------------------

    if remaining > 0:

        print()
        print(
            "WARNING: CRM INVENTORY IS LESS THAN "
            "THE BC REQUIRED QUANTITY."
        )

        print(
            f"BC required: {required_quantity}"
        )

        print(
            f"Inventory shortage: {remaining}"
        )

        print(
            "Retained CRM lines will remain for "
            "normal CRM validation."
        )

    else:

        print()
        print(
            "BC QUANTITY FULLY ALLOCATED."
        )

    print()
    print("=" * 70)
    print(
        "CRM INVENTORY ALLOCATION COMPLETED"
    )
    print("=" * 70)

    record_inventory_report(
        inventory_lines,
        required_quantity
    )

    return retained_lines
def read_inventory_for_order_line(
    page,
    item_no,
    product_title,
    occurrence
):

    rows = find_order_rows_by_identity(
        page,
        item_no,
        product_title
    )

    if occurrence >= len(rows):

        raise Exception(
            f"Order-line occurrence {occurrence} "
            f"not found for:\n"
            f"{item_no} | {product_title}\n"
            f"Matching rows found: {len(rows)}"
        )

    row = rows[occurrence]

    cells = row.locator("td")

    if cells.count() < 5:

        raise Exception(
            f"Unexpected order-row structure for:\n"
            f"{item_no} | {product_title}"
        )

    warehouse_cell = cells.nth(3)

    warehouse_select = warehouse_cell.locator(
        ".el-select"
    ).first

    if warehouse_select.count() == 0:

        raise Exception(
            f"Warehouse selector not found for:\n"
            f"{item_no} | {product_title}"
        )

    try:

        current_warehouse = (
            warehouse_select
            .inner_text()
            .strip()
        )

    except Exception:

        current_warehouse = ""

    print()
    print(
        f"Checking inventory:"
    )

    print(
        f"  Item: {item_no}"
    )

    print(
        f"  Product: {product_title}"
    )

    print(
        f"  Occurrence: {occurrence}"
    )

    print(
        f"  Current warehouse: "
        f"{current_warehouse}"
    )

    if (
        WAREHOUSE_NAME.upper()
        not in current_warehouse.upper()
    ):

        print(
            f"  Selecting warehouse: "
            f"{WAREHOUSE_NAME}"
        )

        warehouse_select.click(
            force=True
        )

        select_suam_stores(
            page,
            warehouse_select
        )

        page.wait_for_timeout(1000)

        # Selecting the warehouse can re-render the
        # entire table. Locate the same occurrence again.
        rows = find_order_rows_by_identity(
            page,
            item_no,
            product_title
        )

        if occurrence >= len(rows):

            raise Exception(
                "CRM line disappeared after "
                "warehouse selection:\n"
                f"{item_no} | {product_title}"
            )

        row = rows[occurrence]

        cells = row.locator("td")

    inventory_text = (
        cells.nth(4)
        .inner_text()
        .strip()
    )

    inventory = parse_inventory_value(
        inventory_text
    )

    print(
        f"  Dealer warehouse inventory: "
        f"{inventory}"
    )

    return inventory
def get_new_order_lines(
    before_records,
    after_records
):

    before_counts = Counter(
        record["key"]
        for record in before_records
    )

    after_seen = Counter()

    new_records = []

    for record in after_records:

        key = record["key"]

        current_occurrence = (
            after_seen[key]
        )

        after_seen[key] += 1

        existing_count = (
            before_counts[key]
        )

        # Any occurrence after the number that existed
        # before the search is a newly created CRM line.
        if current_occurrence >= existing_count:

            record_copy = dict(
                record
            )

            record_copy["occurrence"] = (
                current_occurrence
            )

            new_records.append(
                record_copy
            )

    return new_records
# ============================================================
# MAIN TEST
# ============================================================

def inspect_order_row_values(page, product_title):

    print()
    print("=" * 60)
    print("INSPECTING ORDER ROW VALUES")
    print("=" * 60)
    print(f"Product: {product_title}")

    row = find_order_row(
        page,
        product_title
    )

    cells = row.locator("td")
    cell_count = cells.count()

    print(
        f"Number of cells: {cell_count}"
    )

    for i in range(cell_count):
        try:
            cell_text = cells.nth(i).inner_text().strip()

            print()
            print(
                f"CELL {i + 1}:"
            )
            print(
                repr(cell_text)
            )

        except Exception:
            continue

def remove_zero_inventory_order_lines(page):

    print()
    print("=" * 70)
    print("FINAL ZERO-INVENTORY CLEANUP")
    print("=" * 70)

    records = get_order_line_records(page)

    if not records:
        print("No CRM order lines found.")
        return

    # Build occurrence numbers for duplicate product identities.
    seen = Counter()

    lines_to_check = []

    for record in records:

        key = record["key"]

        occurrence = seen[key]

        seen[key] += 1

        record_copy = dict(record)

        record_copy["occurrence"] = occurrence

        lines_to_check.append(record_copy)

    zero_lines = []

    for line in lines_to_check:

        inventory = read_inventory_for_order_line(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"]
        )

        print(
            f"Checking {line['item_no']} | "
            f"{line['product_title']} | "
            f"Inventory={inventory}"
        )

        if inventory <= 0:

            zero_lines.append(line)

    if not zero_lines:

        print()
        print("No zero-inventory CRM lines remain.")
        return

    # Highest occurrence first so duplicate identities
    # do not shift before deletion.
    zero_lines.sort(
        key=lambda line: (
            line["key"],
            line["occurrence"]
        ),
        reverse=True
    )

    print()
    print(
        f"Zero-inventory lines to delete: "
        f"{len(zero_lines)}"
    )

    for line in zero_lines:

        delete_order_line_occurrence(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"]
        )

    print()
    print("FINAL ZERO-INVENTORY CLEANUP COMPLETED.")
def test_submit_order(page):

    print()
    print("=" * 60)
    print("TESTING SUBMIT")
    print("=" * 60)

    submit_buttons = page.locator(
        'button:has-text("submit")'
    )

    visible_submit_buttons = []

    count = submit_buttons.count()

    print(
        f"Submit button elements found: {count}"
    )

    for i in range(count):
        try:
            button = submit_buttons.nth(i)

            if button.is_visible():
                visible_submit_buttons.append(button)

        except Exception:
            continue

    print(
        f"Visible Submit buttons found: "
        f"{len(visible_submit_buttons)}"
    )

    if len(visible_submit_buttons) == 0:
        raise Exception(
            "Visible Submit button was not found."
        )

    submit_button = visible_submit_buttons[0]

    print(
        f"Submit button text: "
        f"'{submit_button.inner_text().strip()}'"
    )

    submit_button.scroll_into_view_if_needed()

    print()
    print("Clicking Submit...")

    submit_button.click()

    print(
        "Submit button clicked successfully."
    )

    print()
    print(
        "Waiting for CRM response..."
    )

    page.wait_for_timeout(3000)

    print(
        f"Current page URL after Submit: "
        f"{page.url}"
    )

    print()
    print(
        "Visible page text after Submit:"
    )

    try:
        body_text = page.locator("body").inner_text()

        print(
            body_text[-3000:]
        )

    except Exception as e:
        print(
            f"Could not read page text: {e}"
        )
    print()
    print("=" * 60)
    print("TESTING GO COLLECT MONEY")
    print("=" * 60)

    print("Waiting for Go Collect Money screen...")
    page.wait_for_timeout(2000)

    print("Looking for Go Collect Money button...")

    go_collect_buttons = page.locator(
        'button:has-text("Go collect money")'
    )

    visible_go_collect_buttons = []

    go_collect_count = go_collect_buttons.count()

    print(
        f"Go Collect Money button elements found: "
        f"{go_collect_count}"
    )

    for i in range(go_collect_count):
        try:
            button = go_collect_buttons.nth(i)

            if button.is_visible():
                visible_go_collect_buttons.append(button)

        except Exception:
            continue

    print(
        f"Visible Go Collect Money buttons found: "
        f"{len(visible_go_collect_buttons)}"
    )

    if len(visible_go_collect_buttons) == 0:
        raise Exception(
            "Visible Go Collect Money button was not found."
        )

    go_collect_button = visible_go_collect_buttons[0]

    print(
        f"Go Collect Money button text: "
        f"'{go_collect_button.inner_text().strip()}'"
    )

    go_collect_button.scroll_into_view_if_needed()

    print("Clicking Go collect money...")

    go_collect_button.click()

    print(
        "Go collect money button clicked successfully."
    )

    print()
    print("Waiting for payment screen...")

    # The CRM may replace/close the current Playwright page
    # while keeping the browser window open.
    

    try:
        context = page.context

        print(
            f"Playwright currently sees "
            f"{len(context.pages)} page(s)."
        )

        active_pages = [
            p for p in context.pages
            if not p.is_closed()
        ]

        print(
            f"Active Playwright pages found: "
            f"{len(active_pages)}"
        )

        if len(active_pages) == 0:
            raise Exception(
                "No active Playwright page was found "
                "after clicking Go collect money."
            )

        # Use the last active page because the CRM may
        # have opened/replaced the payment page.
        page = active_pages[-1]

        print(
            "Active payment page selected."
        )

        print(
            f"Payment page URL: {page.url}"
        )

    except Exception as e:
        print(
            f"Could not identify the active payment page: {e}"
        )
        raise

    

    print()
    print("Selecting payment method: Bank")

    # ------------------------------------------------------------
    # SELECT PAYMENT METHOD
    # ------------------------------------------------------------

    print()
    print("Looking for Payment Methods select...")

    # ------------------------------------------------------------
    # WAIT FOR PAYMENT METHODS CONTROL
    # ------------------------------------------------------------

    print()
    print("Waiting for Payment Methods control to appear...")

    payment_selects = page.locator(
        "div.el-select.avue-select"
    )

    for attempt in range(30):

        select_count = payment_selects.count()

        print(
            f"Payment UI check {attempt + 1}/30: "
            f"Avue selects={select_count}"
        )

        if select_count >= 4:
            break

        page.wait_for_timeout(1000)

    else:
        raise Exception(
            "Payment Methods control did not appear "
            "within 30 seconds."
        )

    print()
    print("Payment Methods control is now present.")

    # ------------------------------------------------------------
    # SELECT THE 4TH AVUE SELECT
    # ------------------------------------------------------------

    payment_select = payment_selects.nth(3)

    print()
    print(
        "Using the 4th Avue select as Payment Methods."
    )

    print(
        f"Payment select visible: "
        f"{payment_select.is_visible()}"
    )

    if not payment_select.is_visible():
        raise Exception(
            "Payment Methods select is present "
            "but not visible."
        )

    print()
    print("Opening Payment Methods dropdown...")

    payment_select.click()

    print(
        "Payment Methods dropdown opened."
    )

    page.wait_for_timeout(500)

    # ------------------------------------------------------------
    # SELECT BANK
    # ------------------------------------------------------------

    print()
    print("Selecting payment method: Bank")

    bank_option = page.locator(
        '[role="option"]'
    ).filter(
        has_text="bank"
    )

    print(
        f"Bank option elements found: "
        f"{bank_option.count()}"
    )

    if bank_option.count() == 0:

        print()
        print("Available payment options:")

        options = page.locator(
            '[role="option"]'
        )

        for i in range(options.count()):

            try:
                print(
                    f"  - {options.nth(i).inner_text().strip()}"
                )
            except Exception:
                pass

        raise Exception(
            "Bank payment option could not be found."
        )

    bank_option.first.click()

    print(
        "Bank payment method selected successfully."
    )

    page.wait_for_timeout(500)

    

    print()
    print("Looking for payment Submit button...")

    payment_submit_button = page.locator(
        "button.el-button.el-button--primary"
    ).filter(
        has_text="submit"
    )

    print(
        f"Payment Submit buttons found: "
        f"{payment_submit_button.count()}"
    )

    visible_payment_submit_buttons = []

    for i in range(
        payment_submit_button.count()
    ):
        try:
            button = payment_submit_button.nth(i)

            if button.is_visible():
                visible_payment_submit_buttons.append(
                    button
                )

        except Exception:
            continue

    print(
        f"Visible payment Submit buttons found: "
        f"{len(visible_payment_submit_buttons)}"
    )

    if len(visible_payment_submit_buttons) == 0:
        raise Exception(
            "Visible payment Submit button "
            "was not found."
        )

    payment_submit_button = (
        visible_payment_submit_buttons[-1]
    )

    print(
        f"Payment Submit button text: "
        f"'{payment_submit_button.inner_text().strip()}'"
    )

    payment_submit_button.scroll_into_view_if_needed()

    print()
    print("Clicking payment Submit...")

    payment_submit_button.click()

    print(
        "Payment Submit clicked successfully."
    )

    print()
    print("Waiting for payment response...")

    try:
        page.wait_for_load_state(
            "domcontentloaded",
            timeout=5000
        )
    except Exception:
        pass

    page.wait_for_timeout(2000)

    print()
    print("=" * 60)
    print("PAYMENT TEST COMPLETED")
    print("=" * 60)

    try:
        body_text = page.locator(
            "body"
        ).inner_text()

        print()
        print(
            "Visible page text after payment:"
        )

        print(
            body_text[-4000:]
        )

    except Exception as e:
        print(
            f"Could not read page text after "
            f"payment: {e}"
        )
# ============================================================
# EXCEL -> CRM ORDER PROCESSING HELPERS
# ============================================================

PRICE_PREFIX_RULES = {
    "24": 1200,
    "33": 1300,
    "44": 1550,
    "66": 2000,
    "36": 2200,
    "40": 2200,
    "45": 2200,
    "55": 2200,
    "61": 2200,
}


def get_crm_search_text(line):
    """
    Determine what should be entered into the CRM
    'The title of the product' search field.
    """

    description = str(
        line["description"]
    ).strip()

    item_no = str(
        line["item_no"]
    ).strip().upper()

    if not description:
        raise Exception(
            f"Cannot determine CRM search text for "
            f"item {item_no} because the Excel "
            f"description is blank."
        )

    description_upper = description.upper()

    # --------------------------------------------------------
    # FRENCIA
    # --------------------------------------------------------
    #
    # Frencia products are searched using the embedded
    # CRM-style code, for example:
    #
    # KF-028
    # KS-7843S
    # KS-1050
    # WC-793P
    # WC-009
    # WT-02D
    #
    # If a code is embedded in the description, use it.
    # --------------------------------------------------------

    # --------------------------------------------------------
    # ASIAN TOILET / SQUATTING PAN
    # --------------------------------------------------------

    if (
        "ASIAN TOILET" in description_upper
        or "SQUATTING PAN" in description_upper
        or "C.T. PAN" in description_upper
    ):
        return "SQ"

    # --------------------------------------------------------
    # PB-326 / PB-207 / PB-828 PAIRED PRODUCTS
    # --------------------------------------------------------

    if (
        "PB-326" in description_upper
        or "PB-207" in description_upper
        or "PB-828" in description_upper
    ):
        if "PB-326" in description_upper:
            return "PB-326"

        if "PB-207" in description_upper:
            return "PB-207"

        return "PB-828"
    # --------------------------------------------------------
    # OTHER FRENCIA PRODUCTS
    # --------------------------------------------------------

    if "FRENCIA" in description_upper:

        embedded_code = extract_embedded_crm_code(
            description
        )

        if embedded_code:

            if embedded_code in {
                "WC-004T",
                "WC-004P",
                "WC-006T",
                "WC-006P",
                "WC-008T",
                "WC-008P",
                "WC-009T",
                "WC-009P",
                "WC-029P",
                "WC-100P",
                "WC-100T",
            }:

                return embedded_code[:-1]

            return embedded_code

        if item_no:
            return item_no

        raise Exception(
            f"FRENCIA product does not contain a "
            f"recognizable CRM code:\n{description}"
        )

    # --------------------------------------------------------
    # MRP TILE CODES
    # --------------------------------------------------------
    # MRP products use 6 digits after "MRP".
    # The MRP code may appear in the item number or
    # inside the product description.
    #
    # Example:
    # MRP612003Z -> CRM search = 612003
    # MRP612002Y -> CRM search = 612002
    # --------------------------------------------------------

    mrp_code = re.search(
        r"\bMRP(\d{6})(?:[A-Z])?\b",
        description_upper
    )

    if not mrp_code:
        mrp_code = re.search(
            r"\bMRP(\d{6})(?:[A-Z])?\b",
            item_no
        )

    if mrp_code:
        return mrp_code.group(1)

    # --------------------------------------------------------
    # FT / WT TILES
    # --------------------------------------------------------

    if (
        description_upper.startswith("FT ")
        or description_upper.startswith("WT ")
    ):

        tile_code = re.search(
        r"\d{5}",
        description
    )

        if tile_code:
           return tile_code.group(0)

    # --------------------------------------------------------
    # NORMAL PRODUCTS
    # --------------------------------------------------------

    return description

    


def get_crm_price(
    line,
    order_lines=None
):
    """
    Determine the CRM unit price for a normal product.

    Special-product exceptions will be handled separately.
    """

    item_no = str(
        line["item_no"]
    ).strip()

    description = str(
        line["description"]
    ).upper()

    unit_price_excel = float(
    line["unit_price_excel"]
)

    amount_excel = line.get(
        "amount_excel"
    )

    if (
        not unit_price_excel.is_integer()
        and amount_excel is not None
    ):

        excel_price = float(
            amount_excel
        )

    else:

        excel_price = unit_price_excel

    # --------------------------------------------------------
    # SPECIAL PRICE EXCEPTIONS
    # --------------------------------------------------------

    # Frencia and Asian toilet products use
    # the Excel/BC unit price.
    if (
        "FRENCIA" in description
        or "ASIAN TOILET" in description
        or "SQUATTING PAN" in description
        or "C.T. PAN" in description
    ):
        return excel_price

    # --------------------------------------------------------
    # SC-001 FIXED PRICE
    # --------------------------------------------------------

    if item_no == "SC-001":
        return 1000

    # --------------------------------------------------------
    # PAIRED PRODUCT PRICING
    # --------------------------------------------------------

    order_item_nos = {
        str(order_line["item_no"]).strip().upper()
        for order_line in (order_lines or [])
    }

    # WC-006T + WC-006P + SC-001
    # WC-008T + WC-008P + SC-001
    # WC-009T + WC-009P + SC-001
    # SC-001 is not present in the Excel/BC order lines.
    toilet_pairs_with_sc = [
        {"WC-006T", "WC-006P"},
        {"WC-008T", "WC-008P"},
        {"WC-009T", "WC-009P"},
    ]

    if any(
        item_no in pair and pair.issubset(order_item_nos)
        for pair in toilet_pairs_with_sc
    ):
        return (excel_price - 1000) / 2

    # WC-100T + WC-100P
    if item_no in {"WC-100T", "WC-100P"}:
        if {"WC-100T", "WC-100P"}.issubset(order_item_nos):
            return excel_price / 2

    # WC-001W/P/B + matching PB-001W/P/B
    matching_pairs = {
        "WC-001W": "PB-001W",
        "WC-001P": "PB-001P",
        "WC-001B": "PB-001B",
        "WC-004P": "WC-004T",
        "WC-004T": "WC-004P",
        "PB-326P": "PB-326B",
        "PB-326B": "PB-326P",
        "PB-207P": "PB-207B",
        "PB-207B": "PB-207P",
        "PB-828P": "PB-828B",
        "PB-828B": "PB-828P",
    }

    if item_no in matching_pairs:
        matching_item = matching_pairs[item_no]

        if matching_item in order_item_nos:
            return excel_price / 2

    # --------------------------------------------------------
    # STANDARD DIGIT-PREFIX PRICING
    # --------------------------------------------------------

    pricing_code = item_no

    if (
        "FT " in description
        or "WT " in description
    ):

        embedded_code = re.search(
            r"(?<!\d)(?:24|33|44|66|61|40|55|36|45)\d{3,4}(?!\d)",
            description
        )

        if embedded_code:
            pricing_code = embedded_code.group(0)

    for prefix, price in PRICE_PREFIX_RULES.items():

        if pricing_code.startswith(prefix):
            return price

    # --------------------------------------------------------
    # DEFAULT
    # --------------------------------------------------------

    return excel_price


def parse_inventory_value(inventory_text):
    """
    Convert the CRM Dealer warehouse inventory text
    into a numeric quantity.
    """

    text = str(
        inventory_text
    ).strip()

    if not text:
        return 0

    cleaned = (
        text
        .replace(",", "")
        .strip()
    )

    try:
        return float(cleaned)

    except Exception as e:

        raise Exception(
            f"Could not parse CRM inventory value "
            f"'{inventory_text}'."
        ) from e


def print_excel_order_plan(order):
    """
    Display the exact order data that will be sent
    to the CRM.
    """

    print()
    print("=" * 80)
    print(
        f"ORDER PLAN: {order['document_no']}"
    )
    print("=" * 80)

    print(
        f"Sell-to Customer No.: "
        f"{order['customer_no']}"
    )

    print(
        f"Lines: {len(order['lines'])}"
    )

    for index, line in enumerate(
        order["lines"],
        start=1
    ):

        search_text = get_crm_search_text(
            line
        )

        crm_price = get_crm_price(
            line,
            order['lines']
        )

        print()
        print(
            f"LINE {index}"
        )

        print(
            f"  BC Item: "
            f"{line['item_no']}"
        )

        print(
            f"  Description: "
            f"{line['description']}"
        )

        print(
            f"  CRM Search: "
            f"{search_text}"
        )

        print(
            f"  Quantity: "
            f"{line['quantity']}"
        )

        print(
            f"  Excel Price: "
            f"{line['unit_price_excel']}"
        )

        print(
            f"  CRM Price: "
            f"{crm_price}"
        )

        print(
            f"  UOM: "
            f"{line['unit_of_measure']}"
        )

    print("=" * 80)
def main():
     
    if EXCEL_PREVIEW_ONLY:

        orders = load_excel_orders(
            EXCEL_FILE
        )

        print()
        print("=" * 70)
        print("EXCEL ORDER PREVIEW")
        print("=" * 70)

        preview_count = min(
            len(orders),
            10
        )

        for index, order in enumerate(
            list(orders.values())[:preview_count],
            start=1
        ):

            print()
            print(
                f"ORDER {index}:"
            )

            print(
                f"Document No.: "
                f"{order['document_no']}"
            )

            print(
                f"Customer No.: "
                f"{order['customer_no']}"
            )

            print(
                f"Lines: "
                f"{len(order['lines'])}"
            )

            for line in order["lines"]:

                print(
                    f"  {line['item_no']} | "
                    f"{line['description']} | "
                    f"Qty={line['quantity']} | "
                    f"Excel Price={line['unit_price_excel']}"
                )

        print()
        print(
            "EXCEL PREVIEW COMPLETED."
        )

        return

    with sync_playwright() as p:

        browser = None
        context = None

        try:

            print(
                "Opening CRM..."
            )

            browser = p.chromium.launch(
                executable_path=BRAVE_PATH,
                headless=False,
                args=[
                    "--start-maximized",
                ],
            )

            context = browser.new_context(
                no_viewport=True
            )

            page = context.new_page()

            page.goto(
                CRM_URL,
                wait_until="domcontentloaded",
                timeout=60000
            )

            print()
            print(
                "Please log in manually and complete MFA "
                "if required."
            )

            # ------------------------------------------------
            # MANUAL LOGIN
            # ------------------------------------------------

            page.get_by_text(
                "Customer/Sales Management",
                exact=True
            ).wait_for(
                state="visible",
                timeout=300000
            )

            print(
                "LOGIN CONFIRMED"
            )

            # ------------------------------------------------
            # CUSTOMER / SALES MANAGEMENT
            # ------------------------------------------------

            page.get_by_text(
                "Customer/Sales Management",
                exact=True
            ).click(force=True)

            page.wait_for_timeout(1000)

            print(
                "Customer/Sales Management opened."
            )

            # ------------------------------------------------
            # CUSTOMER ORDER MANAGEMENT
            # ------------------------------------------------

            customer_order_management = page.locator(
                'li.el-menu-item[data-track-args-code="custOrderSearch"]'
            )

            customer_order_management.wait_for(
                state="visible",
                timeout=15000
            )

            customer_order_management.click(
                force=True
            )

            page.wait_for_timeout(1200)

            print(
                "Customer Order management opened."
            )

            # ------------------------------------------------
            # CREATE ORDER
            # ------------------------------------------------

            print()
            print("=" * 60)
            print("OPENING CREATE AN ORDER")
            print("=" * 60)

            create_order = page.get_by_role(
                "button",
                name="Create an order",
                exact=True
            )

            create_order.wait_for(
                state="visible",
                timeout=15000
            )

            print("Create an order button found.")

            # The CRM may open the order entry page in a new tab,
            # or it may navigate the existing tab. Support both.
            pages_before = context.pages

            print(
                f"Browser pages before clicking: "
                f"{len(pages_before)}"
            )

            current_page_before = page

            print("Clicking Create an order...")

            create_order.click(force=True)

            print("Create an order clicked.")

            page.wait_for_timeout(1500)

            pages_after = context.pages

            print(
                f"Browser pages after clicking: "
                f"{len(pages_after)}"
            )

            new_pages = [
                p
                for p in pages_after
                if p not in pages_before
            ]

            if new_pages:
                print()
                print("NEW ORDER TAB DETECTED.")

                order_page = new_pages[-1]

                try:
                    order_page.wait_for_load_state(
                        "domcontentloaded",
                        timeout=30000
                    )
                except PlaywrightTimeoutError:
                    print(
                        "Order tab did not finish "
                        "domcontentloaded within timeout; "
                        "continuing because the page exists."
                    )
            else:
                print()
                print("No new tab detected.")
                print(
                    "Checking whether the existing page "
                    "became the order page..."
                )

                order_page = current_page_before
                order_page.wait_for_timeout(1500)

            print()
            print(
                f"Current order page URL: "
                f"{order_page.url}"
            )

            # Verify the actual order-entry page before continuing.
            order_page_indicator = order_page.locator(
                'input[placeholder="Please input Sell customers"]'
            )

            try:
                order_page_indicator.wait_for(
                    state="visible",
                    timeout=15000
                )
                print()
                print("ORDER ENTRY PAGE CONFIRMED.")
            except PlaywrightTimeoutError:
                print()
                print(
                    "Sell customers field was not immediately found."
                )

                order_page.wait_for_timeout(2000)

                try:
                    order_page_indicator.wait_for(
                        state="visible",
                        timeout=10000
                    )
                    print(
                        "ORDER ENTRY PAGE CONFIRMED "
                        "AFTER ADDITIONAL WAIT."
                    )
                except PlaywrightTimeoutError:
                    raise Exception(
                        "Create an order was clicked, but the CRM "
                        "order-entry page could not be confirmed. "
                        f"Current URL: {order_page.url}"
                    )

            print()
            print(
                f"Order page: {order_page.url}"
            )

            # ORDER TYPE
            # ------------------------------------------------

            print()
            print(
                "Selecting Retail order..."
            )

            order_type_item = order_page.locator(
                ".el-form-item"
            ).filter(
                has_text="Order Type:"
            )

            order_type_select = order_type_item.locator(
                ".el-select"
            ).first

            order_type_select.click(
                force=True
            )

            # FIX:
            # This must use order_page, not page.

            order_page.wait_for_timeout(500)

            retail_option = order_page.get_by_text(
                "Retail order",
                exact=True
            )

            retail_visible = None

            for i in range(
                retail_option.count()
            ):

                candidate = retail_option.nth(i)

                try:

                    if candidate.is_visible():

                        retail_visible = candidate
                        break

                except Exception:
                    pass

            if retail_visible is None:

                raise Exception(
                    "Retail order option not found."
                )

            retail_visible.click(
                force=True
            )

            order_page.wait_for_timeout(500)

            print(
                "Order Type selected: Retail order"
            )

            # ------------------------------------------------
            # SELL CUSTOMER
            # ------------------------------------------------

            customer_name = get_next_customer_name()

            print(
                f"Entering Sell customers name: "
                f"{customer_name}"
            )

            sell_customer = order_page.locator(
                'input[placeholder="Please input Sell customers"]'
            )

            sell_customer.wait_for(
                state="visible",
                timeout=15000
            )

            sell_customer.fill(
                customer_name
            )

            print(
                f"Sell customers entered: "
                f"{customer_name}"
            )

            # ------------------------------------------------
            # AUTOMATIC OUTBOUND SHIPMENT
            # ------------------------------------------------

            print()
            print(
                "Checking Automatic outbound shipment..."
            )

            shipment_label = order_page.locator(
                "label.el-checkbox"
            ).filter(
                has_text="Automatic outbound shipment"
            )

            shipment_label.wait_for(
                state="visible",
                timeout=15000
            )

            shipment_checkbox = shipment_label.locator(
                'input.el-checkbox__original[type="checkbox"]'
            )

            visible_checkbox = shipment_label.locator(
                ".el-checkbox__inner"
            )

            initial_state = shipment_checkbox.is_checked()

            print(
                f"Initial checkbox state: {initial_state}"
            )

            if not initial_state:

                visible_checkbox.scroll_into_view_if_needed()

                print(
                    "Clicking Automatic outbound shipment..."
                )

                visible_checkbox.click(
                    force=True
                )

                print(
                    "Automatic outbound shipment clicked."
                )

            else:

                print(
                    "Automatic outbound shipment was already selected."
                )

            # ------------------------------------------------
            # HANDLE POSSIBLE CONFIRMATION DIALOG
            # ------------------------------------------------

            print()
            print(
                "Checking for automatic shipment "
                "confirmation dialog..."
            )

            shipment_confirmed = False

            order_page.wait_for_timeout(700)

            try:

                confirm_buttons = order_page.get_by_role(
                    "button",
                    name="Confirm",
                    exact=True
                )

                visible_confirm_buttons = []

                for i in range(confirm_buttons.count()):

                    button = confirm_buttons.nth(i)

                    try:
                        if button.is_visible():
                            visible_confirm_buttons.append(button)
                    except Exception:
                        continue

                print(
                    f"Visible Confirm buttons found: "
                    f"{len(visible_confirm_buttons)}"
                )

                if visible_confirm_buttons:

                    print(
                        "Shipment confirmation dialog detected."
                    )

                    visible_confirm_buttons[-1].click(
                        force=True
                    )

                    shipment_confirmed = True

                    print(
                        "Automatic outbound shipment confirmed."
                    )

            except Exception as e:

                print(
                    f"Confirmation dialog check produced: {e}"
                )

            if not shipment_confirmed:

                print(
                    "No shipment confirmation dialog detected."
                )

            # ------------------------------------------------
            # FINAL VERIFICATION
            # ------------------------------------------------

            print()
            print(
                "Verifying Automatic outbound shipment state..."
            )

            order_page.wait_for_timeout(1000)

            final_state = shipment_checkbox.is_checked()

            print(
                f"Checkbox state after confirmation: "
                f"{final_state}"
            )

            if not final_state:

                raise Exception(
                    "Automatic outbound shipment checkbox "
                    "is not selected after confirmation."
                )

            print()
            print(
                "Automatic outbound shipment step complete."
            )

            

            # ------------------------------------------------
            # LOAD REAL EXCEL ORDER DATA
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("LOADING REAL EXCEL ORDER DATA")
            print("=" * 70)

            excel_orders = load_excel_orders(
                EXCEL_FILE
            )

            if not excel_orders:
                raise Exception(
                    "No usable Excel order data was found."
                )

            # Flatten all Excel order lines into one CRM batch.
            # This keeps the CRM process fully automated instead
            # of submitting one hard-coded test product.
            excel_lines = []

            for document_no, order in excel_orders.items():

                lines = order["lines"]

                print()
                print(
                    f"Excel Document No. {document_no}: "
                    f"{len(lines)} line(s)"
                )

                for line in lines:

                    excel_lines.append(
                        line
                    )

            print()
            print(
                f"Total Excel orders loaded: "
                f"{len(excel_orders)}"
            )

            print(
                f"Total Excel item lines loaded: "
                f"{len(excel_lines)}"
            )

            if not excel_lines:
                raise Exception(
                    "Excel file was loaded, but no item lines "
                    "are available for CRM entry."
                )

            # ------------------------------------------------
            # ADD ALL EXCEL PRODUCTS TO CRM
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("ADDING EXCEL PRODUCTS TO CRM")
            print("=" * 70)

            

            for index, line in enumerate(
                excel_lines,
                start=1
            ):

                document_no = str(
                    line["document_no"]
                ).strip()

                item_no = str(
                    line["item_no"]
                ).strip()

                description = str(
                    line["description"]
                ).strip()

                quantity = line["quantity"]

                search_text = get_crm_search_text(
                    line
                )
                CURRENT_REPORT_CONTEXT.update(
                {
                    "customer_name": customer_name,
                    "document_no": document_no,
                    "excel_item_no": item_no,
                    "excel_description": description,
                    "excel_quantity": quantity,
                }
            )
                print()
                print("-" * 70)
                print(
                    f"EXCEL LINE {index} / "
                    f"{len(excel_lines)}"
                )
                print(
                    f"Document No.: {document_no}"
                )
                print(
                    f"Item No.: {item_no}"
                )
                print(
                    f"Description: {description}"
                )
                print(
                    f"Quantity: {quantity}"
                )
                print(
                    f"CRM Search: {search_text}"
                )

                # ------------------------------------------------
                # SNAPSHOT ORDER LINES BEFORE PRODUCT SELECTION
                # ------------------------------------------------

                before_order_lines = get_order_line_records(
                    order_page
                )

                print()
                print(
                    f"Existing CRM order lines before search: "
                    f"{len(before_order_lines)}"
                )

                # ------------------------------------------------
                # OPEN PRODUCT WINDOW
                # ------------------------------------------------

                open_product_window(
                    order_page
                )

                # ------------------------------------------------
                # SEARCH CRM
                # ------------------------------------------------

                search_product(
                    order_page,
                    search_text
                )

                # ------------------------------------------------
                # SELECT ALL SEARCH RESULTS
                # ------------------------------------------------

                selected_products = select_all_product_results(
                    order_page
                )

                print()
                print(
                    f"Selected CRM search results: "
                    f"{len(selected_products)}"
                )

                # ------------------------------------------------
                # CONFIRM PRODUCT SELECTION
                # ------------------------------------------------

                confirm_product_selection(
                    order_page
                )

                # ------------------------------------------------
                # SNAPSHOT ORDER LINES AFTER PRODUCT SELECTION
                # ------------------------------------------------

                order_page.wait_for_timeout(1000)

                after_order_lines = get_order_line_records(
                    order_page
                )

                print()
                print(
                    f"CRM order lines after Confirm: "
                    f"{len(after_order_lines)}"
                )

                # ------------------------------------------------
                # IDENTIFY ONLY THE NEWLY CREATED LINES
                # ------------------------------------------------

                new_order_lines = get_new_order_lines(
                    before_order_lines,
                    after_order_lines
                )

                print()
                print(
                    f"New CRM order lines created for this "
                    f"BC item: {len(new_order_lines)}"
                )

                if not new_order_lines:

                    raise Exception(
                        "CRM Confirm completed, but no newly "
                        "created order lines were detected."
                    )

                # ------------------------------------------------
                # ALLOCATE BC QUANTITY ACROSS CRM LINES
                # ------------------------------------------------

                current_order_lines = excel_orders[
                    document_no
                ]["lines"]

                # ------------------------------------------------
                # SPECIAL PAIRED TOILET PRODUCTS
                # WC-004 = WC-004P + WC-004T
                # WC-006 + SC-001
                # WC-008 + SC-001
                # WC-009 + SC-001
                # ------------------------------------------------
                # ------------------------------------------------
                # ASIAN TOILET SPECIAL PAIR
                # SQ-001P + SQ-001T
                # ------------------------------------------------

                paired_model = str(
                    search_text
                ).strip().upper()

                paired_toilet_models = {
                    "WC-004",
                    "WC-006",
                    "WC-008",
                    "WC-009",
                    "WC-029",
                    "WC-100",
                    "WC-5096",
                }

                if paired_model == "SQ":

                    print()
                    print(
                        "=" * 70
                    )
                    print(
                        "SPECIAL ASIAN TOILET PROCESS"
                    )
                    print("=" * 70)

                    asian_p_lines = [
                        line_record
                        for line_record in new_order_lines
                        if "SQ-001P"
                        in str(
                            line_record["product_title"]
                        ).upper()
                    ]

                    asian_t_lines = [
                        line_record
                        for line_record in new_order_lines
                        if "SQ-001T"
                        in str(
                            line_record["product_title"]
                        ).upper()
                    ]

                    print(
                        f"SQ-001P CRM lines found: "
                        f"{len(asian_p_lines)}"
                    )

                    print(
                        f"SQ-001T CRM lines found: "
                        f"{len(asian_t_lines)}"
                    )

                    if not asian_p_lines:

                        raise Exception(
                            "No SQ-001P CRM lines "
                            "were created."
                        )

                    if not asian_t_lines:

                        raise Exception(
                            "No SQ-001T CRM lines "
                            "were created."
                        )

                    # ------------------------------------------------
                    # ASIAN TOILET PRICE = EXCEL PRICE / 2
                    # ------------------------------------------------

                    asian_base_price = get_crm_price(
                        line,
                        current_order_lines
                    )

                    asian_component_price = (
                        asian_base_price / 2
                    )

                    print()
                    print(
                        f"Asian Toilet Excel price: "
                        f"{asian_base_price}"
                    )

                    print(
                        f"SQ-001P price: "
                        f"{asian_component_price}"
                    )

                    print(
                        f"SQ-001T price: "
                        f"{asian_component_price}"
                    )

                    # ------------------------------------------------
                    # DELETE OTHER SQ SEARCH RESULTS
                    # ------------------------------------------------

                    asian_pair_lines = (
                        asian_p_lines
                        + asian_t_lines
                    )

                    asian_extra_lines = [
                        line_record
                        for line_record in new_order_lines
                        if line_record
                        not in asian_pair_lines
                    ]

                    asian_extra_lines.sort(
                        key=lambda line: (
                            line["key"],
                            line["occurrence"]
                        ),
                        reverse=True
                    )

                    print()
                    print(
                        "REMOVING NON-PAIR ASIAN TOILET LINES"
                    )

                    for line_record in asian_extra_lines:

                        delete_order_line_occurrence(
                            order_page,
                            line_record["item_no"],
                            line_record["product_title"],
                            line_record["occurrence"]
                        )

                    # ------------------------------------------------
                    # ALLOCATE SQ-001P
                    # ------------------------------------------------

                    retained_lines = []

                    retained_asian_p_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            asian_p_lines,
                            quantity,
                            asian_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_asian_p_lines
                    )

                    # ------------------------------------------------
                    # ALLOCATE SQ-001T
                    # ------------------------------------------------

                    retained_asian_t_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            asian_t_lines,
                            quantity,
                            asian_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_asian_t_lines
                    )
                elif paired_model in paired_toilet_models:

                    print()
                    print(
                        "=" * 70
                    )
                    print(
                        "SPECIAL PAIRED TOILET PROCESS"
                    )
                    print("=" * 70)

                    # ------------------------------------------------
                    # IDENTIFY P AND T COMPONENT LINES
                    # ------------------------------------------------

                    if paired_model == "WC-029":

                        paired_p_lines = [
                            line_record
                            for line_record in new_order_lines
                            if "P-029P"
                            in str(
                                line_record["product_title"]
                            ).upper()
                        ]

                    elif paired_model == "WC-5096":

                        paired_p_lines = [
                            line_record
                            for line_record in new_order_lines
                            if "WC-5096 P-TRAP CLOSE COUPLE BOX WHITE"
                            in str(
                                line_record["product_title"]
                            ).upper()
                        ]

                    else:

                        paired_p_lines = [
                            line_record
                            for line_record in new_order_lines
                            if (
                                f"{paired_model}P"
                                in str(
                                    line_record["product_title"]
                                ).upper()
                            )
                        ]

                    if paired_model == "WC-029":

                        paired_t_lines = [
                            line_record
                            for line_record in new_order_lines
                            if "P-029/145T"
                            in str(
                                line_record["product_title"]
                            ).upper()
                        ]

                    elif paired_model == "WC-5096":

                        paired_t_lines = [
                            line_record
                            for line_record in new_order_lines
                            if "WC-5096 TANK CLOSE COUPLE BOX WHITE"
                            in str(
                                line_record["product_title"]
                            ).upper()
                        ]

                    else:

                        paired_t_lines = [
                            line_record
                            for line_record in new_order_lines
                            if (
                                f"{paired_model}T"
                                in str(
                                    line_record["product_title"]
                                ).upper()
                            )
                        ]

                    print(
                        f"{paired_model}P CRM lines found: "
                        f"{len(paired_p_lines)}"
                    )

                    print(
                        f"{paired_model}T CRM lines found: "
                        f"{len(paired_t_lines)}"
                    )

                    if not paired_p_lines:

                        raise Exception(
                            f"No {paired_model}P CRM lines "
                            "were created."
                        )

                    if not paired_t_lines:

                        raise Exception(
                            f"No {paired_model}T CRM lines "
                            "were created."
                        )

                    # ------------------------------------------------
                    # SC-001 IS REQUIRED ONLY FOR WC-006 / WC-008 / WC-009
                    # WC-004 = WC-004P + WC-004T ONLY
                    # ------------------------------------------------

                    sc_new_lines = []

                    if paired_model not in {
                        "WC-004",
                        "WC-029",
                        "WC-100",
                        "WC-5096",
                    }:

                        print()
                        print(
                            "Adding required SC-001 companion..."
                        )

                        before_sc_lines = (
                            get_order_line_records(
                                order_page
                            )
                        )

                        open_product_window(
                            order_page
                        )

                        search_product(
                            order_page,
                            "SC-001"
                        )

                        selected_sc_products = (
                            select_all_product_results(
                                order_page
                            )
                        )

                        print(
                            f"Selected SC-001 CRM results: "
                            f"{len(selected_sc_products)}"
                        )

                        confirm_product_selection(
                            order_page
                        )

                        order_page.wait_for_timeout(
                            1000
                        )

                        after_sc_lines = (
                            get_order_line_records(
                                order_page
                            )
                        )

                        sc_new_lines = (
                            get_new_order_lines(
                                before_sc_lines,
                                after_sc_lines
                            )
                        )

                        print(
                            f"New SC-001 CRM lines created: "
                            f"{len(sc_new_lines)}"
                        )

                        if not sc_new_lines:

                            raise Exception(
                                "SC-001 selection was confirmed, "
                                "but no new SC-001 CRM line "
                                "was detected."
                            )

                    # ------------------------------------------------
                    # CALCULATE PAIRED COMPONENT PRICE
                    # WC-004 = EXCEL PRICE / 2
                    # WC-006 / WC-008 / WC-009 =
                    # (EXCEL PRICE - 1000) / 2
                    # ------------------------------------------------

                    paired_base_price = get_crm_price(
                        line,
                        current_order_lines
                    )

                    if paired_model in {
                        "WC-004",
                        "WC-029",
                        "WC-100",
                        "WC-5096",
                    }:

                        paired_component_price = (
                            paired_base_price / 2
                        )

                    else:

                        paired_component_price = (
                            paired_base_price - 1000
                        ) / 2

                    if paired_component_price < 0:

                        raise Exception(
                            f"Invalid paired-product price "
                            f"for {paired_model}: "
                            f"{paired_component_price}"
                        )

                    print()
                    print(
                        f"Paired base price: "
                        f"{paired_base_price}"
                    )

                    print(
                        f"{paired_model}P price: "
                        f"{paired_component_price}"
                    )

                    print(
                        f"{paired_model}T price: "
                        f"{paired_component_price}"
                    )

                    if paired_model not in {
                        "WC-004",
                        "WC-029",
                        "WC-100",
                        "WC-5096",
                    }:

                        print(
                            "SC-001 price: 1000"
                        )

                    # ------------------------------------------------
                    # ALLOCATE P COMPONENT
                    # ------------------------------------------------

                    retained_lines = []

                    retained_p_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            paired_p_lines,
                            quantity,
                            paired_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_p_lines
                    )

                    # ------------------------------------------------
                    # ALLOCATE T COMPONENT
                    # ------------------------------------------------

                    retained_t_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            paired_t_lines,
                            quantity,
                            paired_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_t_lines
                    )

                    # ------------------------------------------------
                    # ALLOCATE SC-001 COMPONENT
                    # WC-004 does NOT use SC-001
                    # ------------------------------------------------

                    if paired_model != "WC-004":

                        retained_sc_lines = (
                            allocate_quantity_across_new_lines(
                                order_page,
                                sc_new_lines,
                                quantity,
                                1000
                            )
                        )

                        retained_lines.extend(
                            retained_sc_lines
                        )

                elif paired_model in {
                    "PB-326",
                    "PB-207",
                    "PB-828",
                }:

                    print()
                    print("=" * 70)
                    print("SPECIAL PB PAIRED PROCESS")
                    print("=" * 70)

                    # ------------------------------------------------
                    # IDENTIFY P AND B COMPONENT LINES
                    # ------------------------------------------------

                    normalized_model = (
                        paired_model
                        .replace("-", "")
                        .strip()
                        .upper()
                    )

                    paired_p_lines = [
                        line_record
                        for line_record in new_order_lines
                        if (
                            f"{normalized_model}P"
                            in "".join(
                                ch
                                for ch in str(
                                    line_record["product_title"]
                                ).upper()
                                if ch.isalnum()
                            )
                        )
                    ]

                    paired_b_lines = [
                        line_record
                        for line_record in new_order_lines
                        if (
                            f"{normalized_model}B"
                            in "".join(
                                ch
                                for ch in str(
                                    line_record["product_title"]
                                ).upper()
                                if ch.isalnum()
                            )
                        )
                    ]

                    if not paired_p_lines:
                        raise Exception(
                            f"Could not find "
                            f"{paired_model}P CRM line."
                        )

                    if not paired_b_lines:
                        raise Exception(
                            f"Could not find "
                            f"{paired_model}B CRM line."
                        )

                    # ------------------------------------------------
                    # REMOVE ALL UNRELATED SEARCH RESULTS
                    # ------------------------------------------------

                    retained_pair_lines = (
                        paired_p_lines
                        + paired_b_lines
                    )

                    for line_record in new_order_lines:

                        if line_record not in retained_pair_lines:

                            delete_product_line(
                                order_page,
                                line_record["product_title"]
                            )

                    # ------------------------------------------------
                    # PRICE
                    # BC / EXCEL PRICE SPLIT BETWEEN P AND B
                    # ------------------------------------------------

                    paired_base_price = get_crm_price(
                        line,
                        current_order_lines
                    )

                    if paired_model == "WC-004":

                        paired_component_price = (
                            paired_base_price / 2
                        )

                    else:

                        paired_component_price = (
                            paired_base_price - 1000
                        ) / 2

                    if paired_component_price < 0:
                        raise Exception(
                            f"Invalid PB paired-product "
                            f"price for {paired_model}: "
                            f"{paired_component_price}"
                        )

                    print(
                        f"Paired base price: "
                        f"{paired_base_price}"
                    )

                    print(
                        f"{paired_model}P price: "
                        f"{paired_component_price}"
                    )

                    print(
                        f"{paired_model}B price: "
                        f"{paired_component_price}"
                    )

                    # ------------------------------------------------
                    # ALLOCATE P COMPONENT
                    # ------------------------------------------------

                    retained_lines = []

                    retained_p_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            paired_p_lines,
                            quantity,
                            paired_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_p_lines
                    )

                    # ------------------------------------------------
                    # ALLOCATE B COMPONENT
                    # ------------------------------------------------

                    retained_b_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            paired_b_lines,
                            quantity,
                            paired_component_price
                        )
                    )

                    retained_lines.extend(
                        retained_b_lines
                    )

                else:
                    # NORMAL PRODUCT ALLOCATION

                    crm_price = get_crm_price(
                        line,
                        current_order_lines
                    )

                    retained_lines = (
                        allocate_quantity_across_new_lines(
                            order_page,
                            new_order_lines,
                            quantity,
                            crm_price
                        )
                    )

                print()
                print(
                    f"Retained CRM lines for BC item: "
                    f"{len(retained_lines)}"
                )

                print(
                    "Product inventory allocation completed."
                )

                

            

            # ------------------------------------------------
            # SHOW CRM ORDER LINES
            # ------------------------------------------------
            remove_zero_inventory_order_lines(
                order_page
            )
            
            print()
            print("=" * 70)
            print("CRM ORDER LINES AFTER PRODUCT ENTRY")
            print("=" * 70)

            print_current_order_lines(
                order_page
            )

            

            # ------------------------------------------------
            # FINAL CRM ORDER INSPECTION
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("FINAL CRM ORDER STATE")
            print("=" * 70)

            print_current_order_lines(
                order_page
            )

            print()
            print("=" * 70)
            print("INSPECTING CALCULATED CRM VALUES")
            print("=" * 70)

            for line in excel_lines:

                description = str(
                    line["description"]
                ).strip()

                

                try:

                    inspect_order_row_values(
                        order_page,
                        description
                    )

                except Exception as inspection_error:

                    print(
                        f"Could not inspect "
                        f"{description}: "
                        f"{inspection_error}"
                    )

            
            # ------------------------------------------------
            # SUBMIT ORDER
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("SUBMITTING EXCEL-DRIVEN CRM ORDER")
            print("=" * 70)

            test_submit_order(
                order_page
            )

            print()
            print("=" * 70)
            print("EXCEL-DRIVEN CRM ORDER COMPLETED")
            print("=" * 70)

            print()
            print(
                "Processed Excel documents: "
                f"{len(excel_orders)}"
            )

            print(
                "Processed Excel item lines: "
                f"{len(excel_lines)}"
            )
            report_path = write_posting_report()

            if report_path:
                print(
                    f"Posting report saved to: "
                    f"{report_path}"
                )
            print()
            print(
                "CRM calculations were left to the CRM."
            )

            print(
                "The CRM page will remain open for inspection."
            )

            wait_for_enter(
                "FINAL CRM PAGE"
            )

        except Exception as e:

            print()
            print("=" * 70)
            print(
                "TEST STOPPED BECAUSE OF AN ERROR"
            )
            print("=" * 70)

            print(
                f"{type(e).__name__}: {e}"
            )

            print()
            print(
                "FULL TRACEBACK:"
            )

            traceback.print_exc()

            print()

            print(
                "The browser will remain open."
            )

            wait_for_enter(
                "ERROR STATE — INSPECT THE CRM PAGE "
                "BEFORE CLOSING IT"
            )

        finally:

            # IMPORTANT:
            #
            # Do NOT close the browser automatically.
            #
            # This is intentionally disabled while we
            # debug the CRM workflow.

            print()

            print(
                "Browser left open intentionally."
            )

            # No browser.close()
            # No context.close()


if __name__ == "__main__":
    main()

