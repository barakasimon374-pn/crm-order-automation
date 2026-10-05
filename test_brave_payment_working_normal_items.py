from collections import Counter
import os
import sys
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import traceback
import re
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import db as _db

print("### RUNNING NORMAL ITEMS VERSION ###")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ============================================================
# CONFIGURATION
# ============================================================

CRM_URL = "https://smdp4cust.twyfordtile.net/#/wel/index"

BRAVE_PATH = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"

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
    Path(__file__).resolve().parent / "customer_name_rotation.txt"
)

# ============================================================
# CRM POSTING REPORT
# ============================================================

POSTING_REPORT_ROWS = []
ZERO_INVENTORY_REPORT_LINES = []
INVENTORY_RESERVATIONS = {}

# Normal zero-inventory CRM lines are kept temporarily and
# deleted after the next Excel item creates its CRM line.
PENDING_ZERO_INVENTORY_DELETIONS = []


def write_posting_report():
    if not POSTING_REPORT_ROWS:
        print()
        print("No posting report rows were collected.")
        return None

    downloads = Path.home() / "Downloads"
    downloads.mkdir(parents=True, exist_ok=True)

    report_time = time.strftime("%Y-%m-%d_%H%M%S")
    report_path = downloads / f"CRM Posting Report - {report_time}.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Posting Report"

    headers = [
        "Date",
        "Customer",
        "Document No.",
        "Excel Item No.",
        "Excel Description",
        "Excel Qty",
        "CRM Item No.",
        "CRM Product",
        "CRM Inventory",
        "Allocated Qty",
        "Inventory Status",
    ]

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    yellow_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
    bright_yellow_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    red_font = Font(color="C00000")
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
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row in POSTING_REPORT_ROWS:
        worksheet.append([
            row.get("posting_date", ""),
            row.get("customer_no", ""),
            row.get("document_no", ""),
            row.get("excel_item_no", ""),
            row.get("excel_description", ""),
            row.get("excel_qty", ""),
            row.get("crm_item_no", ""),
            row.get("crm_product", ""),
            row.get("crm_inventory", ""),
            row.get("allocated_qty", ""),
            row.get("inventory_status", ""),
        ])

        row_number = worksheet.max_row
        status = row.get("inventory_status", "")

        if status in {"ZERO INVENTORY", "INSUFFICIENT INVENTORY"}:
            worksheet.cell(row=row_number, column=5).font = red_font

        if status in {"ENOUGH INVENTORY", "INSUFFICIENT INVENTORY"}:
            for column in (6, 9, 10, 11):
                worksheet.cell(row=row_number, column=column).fill = yellow_fill

        if status == "ENOUGH INVENTORY":
            worksheet.cell(row=row_number, column=5).fill = bright_yellow_fill
            worksheet.cell(row=row_number, column=5).font = Font(color="000000")

    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    worksheet.row_dimensions[1].height = 32
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)

        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)

    workbook.save(report_path)

    print()
    print("=" * 70)
    print("CRM POSTING REPORT CREATED")
    print("=" * 70)
    print(f"Report: {report_path}")
    print(f"Rows: {len(POSTING_REPORT_ROWS)}")

    return report_path


def get_next_customer_name():
    try:
        index = int(
            CUSTOMER_NAME_STATE_FILE.read_text(encoding="utf-8").strip()
        )
    except (FileNotFoundError, ValueError):
        index = 0

    customer_name = CUSTOMER_NAME_SEQUENCE[index % len(CUSTOMER_NAME_SEQUENCE)]
    CUSTOMER_NAME_STATE_FILE.write_text(str(index + 1), encoding="utf-8")

    return customer_name


WAREHOUSE_NAME = "SUAM STORES"
CRM_EXCEL_FILE = os.environ.get("CRM_EXCEL_FILE")
EXCEL_FILE = Path(CRM_EXCEL_FILE) if CRM_EXCEL_FILE else None
EXCEL_PREVIEW_ONLY = False

EXCLUDED_FT_DESCRIPTIONS = (
    "FT 5209 40.5*40.5 12PCS",
)


def is_approved_excel_description(description):
    """
    Approved products:
        ASIAN TOILET  (any description containing this phrase)
        FRENCIA...    (starts with FRENCIA)
        WT ...        (starts with WT)
        FT ...        (starts with FT) — EXCEPT:
                          FT GG...  (filtered out)
                          FT GS...  (filtered out)
                          explicit EXCLUDED_FT_DESCRIPTIONS list

    FT variants that are kept include (but are not limited to):
        FT FGE, FT FGP, FT MR, FT BLO, and all other non-GG/GS FT codes.
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
        # Filter out FT GG and FT GS variants.
        if upper_text.startswith("FT GG") or upper_text.startswith("FT GS"):
            return False

        normalized = " ".join(upper_text.split())

        for excluded in EXCLUDED_FT_DESCRIPTIONS:
            if normalized == " ".join(excluded.upper().split()):
                return False

        # FT FGE, FT FGP, FT MR, FT BLO and all other FT codes pass.
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
        raise FileNotFoundError(f"Excel file was not found:\n{excel_path}")

    print(f"Excel file: {excel_path}")

    workbook = load_workbook(filename=excel_path, read_only=True, data_only=True)
    worksheet = workbook.active
    rows = worksheet.iter_rows(values_only=True)

    try:
        header_row = next(rows)
    except StopIteration:
        raise Exception("The Excel workbook is empty.")

    headers = [normalize_excel_value(value) for value in header_row]

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

    missing_headers = [h for h in required_headers if h not in headers]

    if missing_headers:
        raise Exception(
            "The Excel file is missing required columns:\n"
            + "\n".join(f"- {h}" for h in missing_headers)
        )

    column_index = {header: index for index, header in enumerate(headers)}

    orders = {}
    source_rows = 0
    non_k_rows = 0
    unapproved_rows = 0
    skipped_rows = 0

    for excel_row in rows:
        if not any(value not in (None, "") for value in excel_row):
            continue

        source_rows += 1

        row = {
            header: (excel_row[index] if index < len(excel_row) else None)
            for header, index in column_index.items()
        }

        location_code = normalize_excel_value(row["Location Code"])

        if not location_code.upper().startswith("K-"):
            non_k_rows += 1
            continue

        row_type = normalize_excel_value(row["Type"])

        if row_type != "Item":
            skipped_rows += 1
            continue

        description = normalize_excel_value(row["Description"])

        if not is_approved_excel_description(description):
            unapproved_rows += 1
            continue

        document_no = normalize_excel_value(row["Document No."])
        customer_no = normalize_excel_value(row["Sell-to Customer No."])
        item_no = normalize_excel_value(row["No."])
        quantity = row["Quantity"]
        unit_price = row["Unit Price Incl. VAT"]
        amount = row["Amount"]
        unit_of_measure = normalize_excel_value(row["Unit of Measure Code"])

        if not document_no:
            raise Exception(
                f"Blank Document No. found in Excel row {source_rows + 1}."
            )

        if not customer_no:
            raise Exception(
                f"Blank Sell-to Customer No. found for Document No. {document_no}."
            )

        if not item_no:
            raise Exception(
                f"Blank item number found for Document No. {document_no}."
            )

        if not description:
            raise Exception(
                f"Blank description found for Document No. {document_no}, item {item_no}."
            )

        if quantity in (None, ""):
            raise Exception(
                f"Blank quantity found for Document No. {document_no}, item {item_no}."
            )

        if unit_price in (None, ""):
            raise Exception(
                f"Blank unit price found for Document No. {document_no}, item {item_no}."
            )

        try:
            quantity_value = float(quantity)
        except Exception as e:
            raise Exception(
                f"Invalid quantity '{quantity}' for Document No. {document_no}, "
                f"item {item_no}."
            ) from e

        if quantity_value.is_integer():
            quantity_value = int(quantity_value)

        try:
            unit_price_value = float(unit_price)
        except Exception as e:
            raise Exception(
                f"Invalid unit price '{unit_price}' for Document No. {document_no}, "
                f"item {item_no}."
            ) from e

        if amount in (None, ""):
            amount_value = None
        else:
            try:
                amount_value = float(amount)
            except Exception as e:
                raise Exception(
                    f"Invalid Amount '{amount}' for Document No. {document_no}, "
                    f"item {item_no}."
                ) from e

        if document_no not in orders:
            orders[document_no] = {
                "document_no": document_no,
                "customer_no": customer_no,
                "lines": [],
            }

        existing_customer = orders[document_no]["customer_no"]

        if existing_customer != customer_no:
            raise Exception(
                f"Document No. {document_no} contains multiple customer numbers:\n"
                f"{existing_customer}\n{customer_no}"
            )

        orders[document_no]["lines"].append({
            "document_no": document_no,
            "customer_no": customer_no,
            "item_no": item_no,
            "description": description,
            "quantity": quantity_value,
            "unit_price_excel": unit_price_value,
            "amount_excel": amount_value,
            "unit_of_measure": unit_of_measure,
            "location_code": normalize_excel_value(row["Location Code"]),
        })

    workbook.close()

    print()
    print(f"Source rows read: {source_rows}")
    print(f"Non K- location rows skipped: {non_k_rows}")
    print(f"Non-approved description rows skipped: {unapproved_rows}")
    print(f"Non-item type rows skipped: {skipped_rows}")
    print(f"Usable order rows: {sum(len(order['lines']) for order in orders.values())}")
    print(f"Unique Document Nos.: {len(orders)}")
    print()
    print("EXCEL ORDER LOAD COMPLETED.")

    return orders



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
            "CRM upload mode detected - "
            "ending automation session without waiting for ENTER."
        )
        return

    input("Press ENTER when you are finished inspecting the CRM page...")


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
    # First try the exact product title.
    rows = page.locator("tr").filter(has_text=product_title)
    count = rows.count()

    if count > 0:
        for i in range(count):
            row = rows.nth(i)
            try:
                if row.is_visible():
                    return row
            except Exception:
                pass

    # Fallback: some CRM rows split the product code and description
    # into separate text blocks so the full title won't match as one string.
    print()
    print("Exact order-line text was not matched.")
    print("Inspecting visible order rows for a partial match...")

    visible_rows = page.locator("tr:visible")
    visible_rows_count = visible_rows.count()

    print(f"Visible order rows found: {visible_rows_count}")

    title_parts = product_title.split()
    possible_identifiers = []

    if len(title_parts) > 0:
        possible_identifiers.append(title_parts[0])

    for part in title_parts:
        if part.isdigit() and len(part) >= 5:
            if part not in possible_identifiers:
                possible_identifiers.append(part)

    for i in range(visible_rows_count):
        row = visible_rows.nth(i)

        try:
            row_text = row.inner_text().strip()

            for identifier in possible_identifiers:
                if identifier in row_text:
                    print()
                    print("POSSIBLE ORDER LINE MATCH:")
                    print(row_text)
                    return row

        except Exception:
            continue

    raise Exception(f"Order line not found:\n{product_title}")


# ============================================================
# WAREHOUSE / INVENTORY
# ============================================================

def select_suam_stores(page, warehouse_select):
    """
    Select SUAM STORES from the warehouse dropdown.

    read_inventory_for_order_line() opens the dropdown before calling this
    function, so this function must NOT click the warehouse selector again.
    """
    print()
    print("=" * 60)
    print("SELECTING WAREHOUSE")
    print("=" * 60)
    print(f"Target warehouse: {WAREHOUSE_NAME}")
    print("Looking for the visible SUAM STORES option...")

    page.wait_for_timeout(800)

    # :visible ensures we only target the currently open dropdown.
    visible_options = page.locator(
        "li.el-select-dropdown__item:visible"
    ).filter(has_text=WAREHOUSE_NAME)

    count = visible_options.count()
    print(f"Visible SUAM STORES options found: {count}")

    if count == 0:
        print()
        print("No visible SUAM STORES option found.")
        print()
        print("Currently visible dropdown options:")

        visible_dropdown_options = page.locator(
            "li.el-select-dropdown__item:visible"
        )

        for i in range(visible_dropdown_options.count()):
            try:
                text = visible_dropdown_options.nth(i).inner_text().strip()
                print(f"Visible option {i}: '{text}'")
            except Exception:
                pass

        raise Exception("Visible SUAM STORES dropdown option was not found.")

    option = visible_options.first

    try:
        print()
        print("Actual visible SUAM STORES option found.")
        print("Option HTML:")
        print(option.evaluate("(el) => el.outerHTML")[:3000])
    except Exception:
        pass

    option.wait_for(state="visible", timeout=5000)
    option.scroll_into_view_if_needed()
    page.wait_for_timeout(300)

    try:
        print()
        print("Clicking visible SUAM STORES option...")
        option.click(timeout=5000)
        print("SUAM STORES option clicked.")

    except Exception as e:
        print(f"Normal click failed: {e}")

        try:
            print("Trying force click...")
            option.click(force=True, timeout=5000)
            print("Force click executed.")

        except Exception as e2:
            print(f"Force click failed: {e2}")

            try:
                print("Trying direct DOM click...")
                option.evaluate("(el) => el.click()")
                print("DOM click executed.")

            except Exception as e3:
                raise Exception("All SUAM STORES click attempts failed.") from e3

    # Give Vue/Element Plus time to update the field.
    page.wait_for_timeout(1000)

    try:
        current_value = warehouse_select.inner_text().strip()
    except Exception:
        current_value = ""

    print()
    print(f"Warehouse field now shows: '{current_value}'")

    if WAREHOUSE_NAME.upper() in current_value.upper():
        print()
        print("=" * 60)
        print("WAREHOUSE SELECTED SUCCESSFULLY")
        print("=" * 60)
        print(f"Confirmed warehouse: {WAREHOUSE_NAME}")
        return True

    # Secondary verification directly from the select element.
    try:
        selected_text = (
            warehouse_select
            .locator(".el-select__selected-item, .el-select__placeholder")
            .last
            .inner_text()
            .strip()
        )

        print(f"Selected-item text: '{selected_text}'")

        if WAREHOUSE_NAME.upper() in selected_text.upper():
            print()
            print("=" * 60)
            print("WAREHOUSE SELECTED SUCCESSFULLY")
            print("=" * 60)
            print(f"Confirmed warehouse: {WAREHOUSE_NAME}")
            return True

    except Exception as e:
        print(f"Secondary verification failed: {e}")

    raise Exception(f"Could not select {WAREHOUSE_NAME}.")



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

            print(f"Selector '{selector}' found {count} element(s).")

            for i in range(count):
                button = buttons.nth(i)

                try:
                    if not button.is_visible():
                        continue

                    print(f"Visible Add Product button found using: {selector}")

                    try:
                        print(f"Button text: '{button.inner_text().strip()}'")
                    except Exception:
                        pass

                    button.scroll_into_view_if_needed()
                    page.wait_for_timeout(300)
                    button.click(force=True)
                    page.wait_for_timeout(1000)

                    print("Add a product button clicked successfully.")
                    return

                except Exception as e:
                    print(f"Could not click candidate {i}: {e}")

        except Exception as e:
            print(f"Selector failed: {selector}")
            print(e)

    print()
    print("=" * 50)
    print("ADD PRODUCT BUTTON NOT FOUND")
    print("=" * 50)
    print()
    print("Visible buttons currently on the order page:")

    try:
        visible_buttons = page.locator("button:visible")
        count = visible_buttons.count()

        print(f"Visible buttons found: {count}")

        for i in range(count):
            try:
                button = visible_buttons.nth(i)
                text = button.inner_text().strip()
                if text:
                    print(f"BUTTON {i}: '{text}'")
            except Exception:
                continue

    except Exception as e:
        print(f"Could not inspect visible buttons: {e}")

    print()
    print("The CRM page will remain open for inspection.")

    raise Exception("Could not locate the Add a product button.")


def search_product(page, search_text):
    print(f"Searching product: {search_text}")
    print("Finding 'The title of the product' search field...")

    search_input = page.locator(
        'input[placeholder="Please input The title of the product"]'
    ).first

    search_input.wait_for(state="visible", timeout=15000)
    print("Correct title search field found.")

    search_input.fill(search_text)

    search_button = page.get_by_role("button", name="Search", exact=True)
    print("Clicking Search...")
    search_button.click(force=True)
    print("Search submitted. Waiting for CRM results table to refresh...")

    result_rows = page.locator("tr:visible")

    before_signature = "\n".join(
        " ".join(text.split()) for text in result_rows.all_inner_texts()
    )

    results_ready = False
    previous_signature = None
    stable_checks = 0

    for attempt in range(30):
        try:
            current_rows = page.locator("tr:visible")
            current_texts = current_rows.all_inner_texts()
            current_signature = "\n".join(
                " ".join(text.split()) for text in current_texts
            )

            visible_result_checkboxes = page.locator("tr:visible .el-checkbox__inner")
            checkbox_count = visible_count(visible_result_checkboxes)

            if (
                current_signature
                and current_signature != before_signature
                and checkbox_count > 0
            ):
                if current_signature == previous_signature:
                    stable_checks += 1
                else:
                    stable_checks = 1

                previous_signature = current_signature

                if stable_checks >= 2:
                    results_ready = True
                    print(f"CRM search results ready after {attempt + 1} check(s).")
                    break
            else:
                previous_signature = None
                stable_checks = 0

        except Exception:
            pass

        page.wait_for_timeout(500)

    if not results_ready:
        raise Exception(
            f"CRM search results did not finish refreshing "
            f"for '{search_text}' within 15 seconds."
        )

    print(f"Search completed for: {search_text}")


# ============================================================
# PRODUCT SELECTION
# ============================================================

def extract_embedded_crm_code(text):
    """
    Extract a CRM-style product code such as KF-028, KS-7843S, WC-009, WT-02D.
    Returns the first matching code, or an empty string.
    """
    text = str(text).upper()
    match = re.search(r"\b[A-Z]{2}-\d{2,5}[A-Z]*\b", text)

    if match:
        return match.group(0)

    return ""


def select_all_product_results(page, search_text):
    print()
    print("=" * 60)
    print("SELECTING ALL CRM PRODUCT RESULTS")
    print("=" * 60)

    rows = page.locator("tr:visible")
    row_count = rows.count()

    print(f"Visible product rows found: {row_count}")

    selected_products = []

    for i in range(row_count):
        row = rows.nth(i)

        try:
            cells = row.locator("td")

            checkbox_input = row.locator('input[type="checkbox"]').first
            visible_checkbox = row.locator(".el-checkbox__inner").first

            if visible_checkbox.count() == 0 and checkbox_input.count() == 0:
                continue

            item_no = ""
            product_title = ""

            if cells.count() >= 1:
                item_no = " ".join(cells.nth(0).inner_text().split()).strip()

            if cells.count() >= 2:
                product_title = " ".join(cells.nth(1).inner_text().split()).strip()

            row_text = " ".join(row.inner_text().split()).strip()

            if (
                "THE TITLE OF THE PRODUCT" in row_text.upper()
                and "STANDARD PACKING UNITS" in row_text.upper()
            ):
                continue

            normalized_search = "".join(
                ch for ch in str(search_text).upper() if ch.isalnum()
            )

            normalized_row = "".join(
                ch for ch in row_text.upper() if ch.isalnum()
            )

            if normalized_search and normalized_search not in normalized_row:
                continue

            if not product_title:
                product_title = row_text

            print()
            print(f"PRODUCT RESULT {len(selected_products) + 1}:")
            print(f"  Item No.: {item_no}")
            print(f"  Product: {product_title}")
            print(f"  Row: {row_text}")

            try:
                if checkbox_input.count() > 0:
                    already_checked = checkbox_input.is_checked()
                else:
                    already_checked = "is-checked" in (
                        visible_checkbox.locator("..").get_attribute("class") or ""
                    )
            except Exception:
                already_checked = False

            print(f"  Already selected: {already_checked}")

            if not already_checked:
                if visible_checkbox.count() > 0:
                    visible_checkbox.scroll_into_view_if_needed()
                    visible_checkbox.click(force=True)
                elif checkbox_input.count() > 0:
                    checkbox_input.evaluate("(el) => el.click()")
                else:
                    raise Exception(
                        "No usable checkbox was found for the CRM product row."
                    )

            page.wait_for_timeout(200)

            final_state = False

            for _ in range(10):
                try:
                    if visible_checkbox.count() > 0:
                        parent_class = (
                            visible_checkbox.locator("..").get_attribute("class") or ""
                        )
                        if "is-checked" in parent_class:
                            final_state = True
                            break

                    if checkbox_input.count() > 0 and checkbox_input.is_checked():
                        final_state = True
                        break

                except Exception:
                    pass

                page.wait_for_timeout(200)

            if not final_state:
                raise Exception(
                    f"Product checkbox could not be selected:\n{row_text}"
                )

            selected_products.append({
                "item_no": item_no,
                "product_title": product_title,
            })

        except Exception as e:
            raise Exception(
                f"Could not process product result row {i}: {e}"
            ) from e

    if not selected_products:
        raise Exception(
            "No selectable CRM product results were found after the search."
        )

    print()
    print(f"Total CRM product results selected: {len(selected_products)}")
    print()
    print("ALL CRM PRODUCT RESULTS SELECTED.")

    return selected_products


def find_surviving_order_page(original_page, pages_before, timeout_seconds=15):
    """
    After CRM Confirm, the page containing the product-selection dialog may
    close while the actual order-entry page survives or is recreated.

    Return the surviving page that contains the Sell customers field.
    """
    context = original_page.context
    order_page_indicator_selector = (
        'input[placeholder="Please input Sell customers"]'
    )
    deadline = time.time() + timeout_seconds

    while time.time() < deadline:
        try:
            active_pages = [p for p in context.pages if not p.is_closed()]
        except Exception:
            active_pages = []

        candidates = []

        if original_page in active_pages and not original_page.is_closed():
            candidates.append(original_page)

        new_pages = [c for c in active_pages if c not in pages_before]

        for candidate in reversed(new_pages):
            if candidate not in candidates:
                candidates.append(candidate)

        for candidate in reversed(active_pages):
            if candidate not in candidates:
                candidates.append(candidate)

        for candidate in candidates:
            try:
                indicator = candidate.locator(order_page_indicator_selector).first

                if indicator.count() > 0 and indicator.is_visible():
                    return candidate

            except Exception:
                continue

        time.sleep(0.25)

    return None


def confirm_product_selection(page):
    print()
    print("Confirming product selection...")

    confirm_buttons = page.get_by_role("button", name="Confirm", exact=True)
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
            "Visible Confirm button was not found after product selection."
        )

    visible_confirm.scroll_into_view_if_needed()

    # Capture the current browser pages before Confirm.
    # The CRM can close the current order page and recreate another page.
    pages_before = list(page.context.pages)
    confirm_click_error = None

    try:
        visible_confirm.click(force=True)
    except Exception as confirm_error:
        confirm_click_error = confirm_error
        print()
        print(f"Confirm click reported an exception: {confirm_error}")
        print("Checking for a surviving CRM order page...")

    surviving_page = find_surviving_order_page(page, pages_before, timeout_seconds=15)

    try:
        active_pages = [p for p in page.context.pages if not p.is_closed()]

        print()
        print(f"Active Playwright pages after Confirm: {len(active_pages)}")

        for page_index, active_page in enumerate(active_pages):
            print(f"ACTIVE PAGE {page_index}: {active_page.url}")

    except Exception as page_check_error:
        print(f"Could not inspect active pages after Confirm: {page_check_error}")

    if surviving_page is None:
        if confirm_click_error is not None:
            raise confirm_click_error

        raise Exception(
            "CRM Confirm did not leave a usable order-entry page. "
            "The original page may have closed and no surviving page "
            "containing the Sell customers field was detected."
        )

    if surviving_page is not page:
        print()
        print("CRM Confirm changed the active order page.")
        print(f"Original page closed/unusable: {page.is_closed()}")
        print(f"Surviving order page URL: {surviving_page.url}")

    try:
        surviving_page.wait_for_timeout(1000)
    except Exception:
        pass

    print()
    print("Product selection confirmed.")

    return surviving_page


# ============================================================
# MENU / DELETE
# ============================================================

def inspect_menu(page, product_title):
    print()
    print("Inspecting Menu for:")
    print(product_title)

    row = find_order_row(page, product_title)
    cells = row.locator("td")
    cell_count = cells.count()

    print(f"Number of cells: {cell_count}")

    if cell_count == 0:
        raise Exception(f"No cells found for {product_title}")

    # Menu is the last column.
    menu_cell = cells.nth(cell_count - 1)

    try:
        menu_text = menu_cell.inner_text().strip()
    except Exception:
        menu_text = ""

    print(f"Menu cell text: '{menu_text}'")

    delete_text = menu_cell.get_by_text("Delete", exact=True)
    delete_count = 0

    for i in range(delete_text.count()):
        try:
            if delete_text.nth(i).is_visible():
                delete_count += 1
        except Exception:
            pass

    print(f"Delete elements found: {delete_count}")

    if delete_count > 0:
        print(f"DELETE AVAILABLE for: {product_title}")
    else:
        print(f"DELETE NOT AVAILABLE for: {product_title}")

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

    info = inspect_menu(page, product_title)

    if info["delete_count"] == 0:
        print()
        print("DELETE IS NOT AVAILABLE.")
        print("The CRM does not currently provide a Delete option for this line.")
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
        print("Delete was detected but no visible Delete element could be clicked.")
        return False

    print()
    print("Delete option is available.")

    visible_delete.scroll_into_view_if_needed()
    visible_delete.click(force=True)
    page.wait_for_timeout(500)

    print("Checking for deletion confirmation...")

    confirmation_found = False

    try:
        confirm_buttons = page.get_by_role("button", name="Confirm", exact=True)

        for i in range(confirm_buttons.count()):
            button = confirm_buttons.nth(i)

            try:
                if button.is_visible():
                    print("Deletion confirmation dialog detected.")
                    button.click(force=True)
                    confirmation_found = True
                    print("Deletion confirmed.")
                    break
            except Exception:
                continue

    except Exception as e:
        print(f"Confirmation check produced: {e}")

    if not confirmation_found:
        print("No deletion confirmation dialog detected.")

    print("Waiting for product line to disappear...")
    page.wait_for_timeout(1000)

    try:
        page.locator("tr").filter(has_text=product_title).first.wait_for(
            state="detached", timeout=10000
        )
        print()
        print(f"Successfully deleted: {product_title}")
        return True

    except PlaywrightTimeoutError:
        remaining = page.locator("tr").filter(has_text=product_title).count()

        if remaining == 0:
            print()
            print(f"Successfully deleted: {product_title}")
            return True

        print()
        print(f"WARNING: Could not verify deletion of {product_title}")
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
            print(f"ROW {i}:")
            print(text)
            found += 1

        except Exception:
            continue

    print()
    print(f"Visible table rows inspected: {found}")


def normalize_order_line_text(value):
    return " ".join(str(value).split()).strip()


def get_order_line_records(page):
    rows = page.locator("tr:visible")
    records = []

    for i in range(rows.count()):
        row = rows.nth(i)

        try:
            cells = row.locator("td")

            if cells.count() < 5:
                continue

            item_no = normalize_order_line_text(cells.nth(0).inner_text())
            product_title = normalize_order_line_text(cells.nth(1).inner_text())

            if not item_no or not product_title:
                continue

            records.append({
                "item_no": item_no,
                "product_title": product_title,
                "key": (item_no.upper(), product_title.upper()),
            })

        except Exception:
            continue

    return records


def find_order_rows_by_identity(page, item_no, product_title):
    target_item = normalize_order_line_text(item_no).upper()
    target_title = normalize_order_line_text(product_title).upper()

    rows = page.locator("tr")
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

            if current_item == target_item and current_title == target_title:
                matches.append(row)

        except Exception:
            continue

    return matches


def set_order_line_quantity(page, item_no, product_title, occurrence, quantity):
    rows = find_order_rows_by_identity(page, item_no, product_title)

    if occurrence >= len(rows):
        raise Exception(
            f"Could not find CRM line occurrence {occurrence} for:\n"
            f"{item_no} | {product_title}"
        )

    row = rows[occurrence]
    quantity_input = row.locator('input[type="number"]').first

    if quantity_input.count() == 0:
        raise Exception(
            f"Quantity input not found for:\n{item_no} | {product_title}"
        )

    quantity_input.scroll_into_view_if_needed()
    quantity_input.fill(str(quantity))

    entered = quantity_input.input_value()
    print(f"  CRM quantity set to: {entered}")

    try:
        entered_numeric = float(entered)
        expected_numeric = float(quantity)
    except Exception:
        raise Exception(
            f"Could not verify quantity for:\n{item_no} | {product_title}"
        )

    if entered_numeric != expected_numeric:
        raise Exception(
            f"Quantity mismatch for:\n{item_no} | {product_title}\n"
            f"Expected: {quantity}\nFound: {entered}"
        )


def set_order_line_unit_price(page, item_no, product_title, occurrence, unit_price):
    rows = find_order_rows_by_identity(page, item_no, product_title)

    if occurrence >= len(rows):
        raise Exception(
            f"Could not find CRM line occurrence {occurrence} for:\n"
            f"{item_no} | {product_title}"
        )

    row = rows[occurrence]
    price_input = row.locator(
        'input[placeholder="Please input unit price"]'
    ).first

    if price_input.count() == 0:
        raise Exception(
            f"Unit price input not found for:\n{item_no} | {product_title}"
        )

    price_input.scroll_into_view_if_needed()
    price_input.fill(str(unit_price))
    page.wait_for_timeout(500)

    # CRM may re-render the row after the price changes — re-find it.
    rows = find_order_rows_by_identity(page, item_no, product_title)

    if occurrence >= len(rows):
        raise Exception(
            f"Could not re-find CRM line after setting unit price for:\n"
            f"{item_no} | {product_title}"
        )

    row = rows[occurrence]
    price_input = row.locator(
        'input[placeholder="Please input unit price"]'
    ).first
    price_input.wait_for(state="visible", timeout=10000)

    entered = price_input.input_value()
    print(f"  CRM unit price set to: {entered}")

    try:
        entered_numeric = float(entered)
        expected_numeric = float(unit_price)
    except Exception:
        raise Exception(
            f"Could not verify unit price for:\n{item_no} | {product_title}"
        )

    if entered_numeric != expected_numeric:
        raise Exception(
            f"Unit price mismatch for:\n{item_no} | {product_title}\n"
            f"Expected: {unit_price}\nFound: {entered}"
        )


def delete_order_line_occurrence(page, item_no, product_title, occurrence):
    rows = find_order_rows_by_identity(page, item_no, product_title)

    if occurrence >= len(rows):
        print(
            f"Line occurrence {occurrence} no longer exists: "
            f"{item_no} | {product_title}"
        )
        return False

    row = rows[occurrence]
    cells = row.locator("td")
    menu_cell = cells.nth(cells.count() - 1)

    delete_button = menu_cell.get_by_text("Delete", exact=True)
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
        print("DELETE NOT AVAILABLE")
        print(f"Line: {item_no} | {product_title}")
        return False

    print()
    print("Deleting CRM line:")
    print(f"  {item_no} | {product_title}")
    print(f"  Occurrence: {occurrence}")

    visible_delete.scroll_into_view_if_needed()
    visible_delete.click(force=True)
    page.wait_for_timeout(400)

    # Confirm deletion dialog if it appears.
    confirm_buttons = page.get_by_role("button", name="Confirm", exact=True)

    for i in range(confirm_buttons.count()):
        button = confirm_buttons.nth(i)
        try:
            if button.is_visible():
                button.click(force=True)
                print("Deletion confirmed.")
                break
        except Exception:
            continue

    page.wait_for_timeout(800)

    remaining_rows = find_order_rows_by_identity(page, item_no, product_title)

    if len(remaining_rows) < len(rows):
        print("CRM line deleted successfully.")
        return True

    print("CRM line could not be verified as deleted.")
    return False


def allocate_quantity_across_new_lines(page, new_lines, required_quantity, unit_price):
    print()
    print("=" * 70)
    print("ALLOCATING BC QUANTITY ACROSS CRM LINES")
    print("=" * 70)

    required_quantity = float(required_quantity)

    if required_quantity <= 0:
        raise Exception(f"Invalid required quantity: {required_quantity}")

    if not new_lines:
        raise Exception(
            "No newly created CRM lines were found for the current BC item."
        )

    print()
    print(f"BC required quantity: {required_quantity}")
    print(f"New CRM lines to inspect: {len(new_lines)}")

    # --------------------------------------------------------
    # READ INVENTORY FROM EVERY NEW CRM LINE
    # --------------------------------------------------------

    inventory_lines = []

    for line in new_lines:
        inventory = read_inventory_for_order_line(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"],
        )

        line_copy = dict(line)

        inventory_key = str(line["item_no"]).strip().upper()
        already_reserved = INVENTORY_RESERVATIONS.get(inventory_key, 0)
        available_inventory = max(0, float(inventory) - float(already_reserved))

        line_copy["inventory"] = available_inventory
        line_copy["allocated_quantity"] = 0
        line_copy["unit_price"] = float(unit_price)

        inventory_lines.append(line_copy)

    # --------------------------------------------------------
    # SEPARATE ZERO-STOCK AND POSITIVE-STOCK LINES
    # --------------------------------------------------------

    zero_lines = [line for line in inventory_lines if line["inventory"] <= 0]
    positive_lines = [line for line in inventory_lines if line["inventory"] > 0]

    ZERO_INVENTORY_REPORT_LINES.extend(dict(line) for line in zero_lines)

    print()
    print(f"Zero-inventory lines: {len(zero_lines)}")
    print(f"Positive-inventory lines: {len(positive_lines)}")

    # --------------------------------------------------------
    # HIGHEST INVENTORY FIRST
    # --------------------------------------------------------

    positive_lines.sort(key=lambda line: line["inventory"], reverse=True)

    # --------------------------------------------------------
    # ALLOCATE ONE SHARED BC QUANTITY
    # --------------------------------------------------------

    remaining = required_quantity
    retained_lines = []
    excess_lines = []

    for line in positive_lines:
        if remaining > 0:
            quantity_to_use = min(line["inventory"], remaining)
            line["allocated_quantity"] = quantity_to_use
            retained_lines.append(line)
            remaining -= quantity_to_use
        else:
            excess_lines.append(line)

    print()
    print(
        f"Total positive inventory available: "
        f"{sum(line['inventory'] for line in positive_lines)}"
    )
    print(f"Remaining BC quantity after allocation: {remaining}")

    # --------------------------------------------------------
    # SET QUANTITY AND PRICE ON RETAINED LINES
    # --------------------------------------------------------

    print()
    print("SETTING QUANTITIES ON RETAINED CRM LINES")

    for line in retained_lines:
        print()
        print(f"Item: {line['item_no']}")
        print(f"Product: {line['product_title']}")
        print(f"Inventory: {line['inventory']}")
        print(f"Allocated quantity: {line['allocated_quantity']}")

        set_order_line_quantity(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"],
            line["allocated_quantity"],
        )

        set_order_line_unit_price(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"],
            unit_price,
        )

    # --------------------------------------------------------
    # QUEUE ZERO-INVENTORY LINES FOR NEXT-ITEM DELETION
    # --------------------------------------------------------

    if zero_lines:
        print()
        print("QUEUING ZERO-INVENTORY CRM LINES FOR NEXT-ITEM DELETION")

        for line in zero_lines:
            PENDING_ZERO_INVENTORY_DELETIONS.append({
                "item_no": str(line["item_no"]).strip(),
                "product_title": str(line["product_title"]).strip(),
                "occurrence": int(line["occurrence"]),
            })

            print(
                f"Queued zero-inventory line: "
                f"{line['item_no']} | {line['product_title']} | "
                f"Occurrence={line['occurrence']}"
            )

    # --------------------------------------------------------
    # DELETE UNUSED POSITIVE-INVENTORY LINES IMMEDIATELY.
    # Zero-inventory lines are deliberately NOT deleted here.
    # --------------------------------------------------------

    if excess_lines:
        excess_lines.sort(
            key=lambda line: (line["key"], line["occurrence"]), reverse=True
        )

        print()
        print("REMOVING UNUSED POSITIVE-INVENTORY CRM LINES")

        for line in excess_lines:
            delete_order_line_occurrence(
                page,
                line["item_no"],
                line["product_title"],
                line["occurrence"],
            )

    # --------------------------------------------------------
    # SHORTFALL REPORT
    # --------------------------------------------------------

    if remaining > 0:
        print()
        print("WARNING: CRM INVENTORY IS LESS THAN THE BC REQUIRED QUANTITY.")
        print(f"BC required: {required_quantity}")
        print(f"Inventory shortage: {remaining}")
        print("Retained CRM lines will remain for normal CRM validation.")
    else:
        print()
        print("BC QUANTITY FULLY ALLOCATED.")

    print()
    print("=" * 70)
    print("CRM INVENTORY ALLOCATION COMPLETED")
    print("=" * 70)

    # --------------------------------------------------------
    # RESERVE INVENTORY USED BY THIS EXCEL BATCH
    # --------------------------------------------------------

    for line in retained_lines:
        inventory_key = str(line["item_no"]).strip().upper()
        allocated = float(line.get("allocated_quantity", 0))
        INVENTORY_RESERVATIONS[inventory_key] = (
            INVENTORY_RESERVATIONS.get(inventory_key, 0) + allocated
        )

    return retained_lines


def read_inventory_for_order_line(page, item_no, product_title, occurrence):
    rows = find_order_rows_by_identity(page, item_no, product_title)

    if occurrence >= len(rows):
        raise Exception(
            f"Order-line occurrence {occurrence} not found for:\n"
            f"{item_no} | {product_title}\n"
            f"Matching rows found: {len(rows)}"
        )

    row = rows[occurrence]
    cells = row.locator("td")

    if cells.count() < 5:
        raise Exception(
            f"Unexpected order-row structure for:\n{item_no} | {product_title}"
        )

    warehouse_cell = cells.nth(3)
    warehouse_select = warehouse_cell.locator(".el-select").first

    if warehouse_select.count() == 0:
        raise Exception(
            f"Warehouse selector not found for:\n{item_no} | {product_title}"
        )

    try:
        current_warehouse = warehouse_select.inner_text().strip()
    except Exception:
        current_warehouse = ""

    print()
    print("Checking inventory:")
    print(f"  Item: {item_no}")
    print(f"  Product: {product_title}")
    print(f"  Occurrence: {occurrence}")
    print(f"  Current warehouse: {current_warehouse}")

    if WAREHOUSE_NAME.upper() not in current_warehouse.upper():
        print(f"  Selecting warehouse: {WAREHOUSE_NAME}")

        warehouse_select.click(force=True)
        select_suam_stores(page, warehouse_select)
        page.wait_for_timeout(1000)

        # Selecting the warehouse can re-render the entire table.
        rows = find_order_rows_by_identity(page, item_no, product_title)

        if occurrence >= len(rows):
            raise Exception(
                f"CRM line disappeared after warehouse selection:\n"
                f"{item_no} | {product_title}"
            )

        row = rows[occurrence]
        cells = row.locator("td")

    inventory_text = cells.nth(4).inner_text().strip()
    inventory = parse_inventory_value(inventory_text)

    print(f"  Dealer warehouse inventory: {inventory}")

    return inventory


def get_new_order_lines(before_records, after_records):
    before_counts = Counter(record["key"] for record in before_records)
    after_seen = Counter()
    new_records = []

    for record in after_records:
        key = record["key"]
        current_occurrence = after_seen[key]
        after_seen[key] += 1

        existing_count = before_counts[key]

        # Any occurrence after the number that existed before the search
        # is a newly created CRM line.
        if current_occurrence >= existing_count:
            record_copy = dict(record)
            record_copy["occurrence"] = current_occurrence
            new_records.append(record_copy)

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

    row = find_order_row(page, product_title)
    cells = row.locator("td")
    cell_count = cells.count()

    print(f"Number of cells: {cell_count}")

    for i in range(cell_count):
        try:
            cell_text = cells.nth(i).inner_text().strip()
            print()
            print(f"CELL {i + 1}:")
            print(repr(cell_text))
        except Exception:
            continue


def remove_zero_inventory_order_lines(page):
    print()
    print("=" * 70)
    print("FINAL ZERO-INVENTORY CLEANUP (SAFETY NET)")
    print("=" * 70)

    records = get_order_line_records(page)

    if not records:
        print("No CRM order lines found.")
        return

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
            line["occurrence"],
        )

        print(
            f"Checking {line['item_no']} | "
            f"{line['product_title']} | Inventory={inventory}"
        )

        if inventory <= 0:
            zero_lines.append(line)

    if not zero_lines:
        print()
        print("No zero-inventory CRM lines remain.")
        return

    # Highest occurrence first so duplicate identities don't shift before deletion.
    zero_lines.sort(
        key=lambda line: (line["key"], line["occurrence"]), reverse=True
    )

    print()
    print(f"Zero-inventory lines to delete: {len(zero_lines)}")

    for line in zero_lines:
        delete_order_line_occurrence(
            page,
            line["item_no"],
            line["product_title"],
            line["occurrence"],
        )

    print()
    print("FINAL ZERO-INVENTORY CLEANUP SAFETY NET COMPLETED.")


def test_submit_order(page):
    print()
    print("=" * 60)
    print("TESTING SUBMIT")
    print("=" * 60)

    submit_buttons = page.locator('button:has-text("submit")')
    visible_submit_buttons = []

    count = submit_buttons.count()
    print(f"Submit button elements found: {count}")

    for i in range(count):
        try:
            button = submit_buttons.nth(i)
            if button.is_visible():
                visible_submit_buttons.append(button)
        except Exception:
            continue

    print(f"Visible Submit buttons found: {len(visible_submit_buttons)}")

    if len(visible_submit_buttons) == 0:
        raise Exception("Visible Submit button was not found.")

    submit_button = visible_submit_buttons[0]
    print(f"Submit button text: '{submit_button.inner_text().strip()}'")

    submit_button.scroll_into_view_if_needed()
    print()
    print("Clicking Submit...")

    submit_button.click()
    print("Submit button clicked successfully.")

    print()
    print("Waiting for CRM response...")

    page.wait_for_timeout(3000)

    print(f"Current page URL after Submit: {page.url}")
    print()
    print("Visible page text after Submit:")

    try:
        body_text = page.locator("body").inner_text()
        print(body_text[-3000:])
    except Exception as e:
        print(f"Could not read page text: {e}")

    print()
    print("=" * 60)
    print("TESTING GO COLLECT MONEY")
    print("=" * 60)

    print("Waiting for Go Collect Money screen...")
    page.wait_for_timeout(2000)

    print("Looking for Go Collect Money button...")

    go_collect_buttons = page.locator('button:has-text("Go collect money")')
    visible_go_collect_buttons = []

    go_collect_count = go_collect_buttons.count()
    print(f"Go Collect Money button elements found: {go_collect_count}")

    for i in range(go_collect_count):
        try:
            button = go_collect_buttons.nth(i)
            if button.is_visible():
                visible_go_collect_buttons.append(button)
        except Exception:
            continue

    print(f"Visible Go Collect Money buttons found: {len(visible_go_collect_buttons)}")

    if len(visible_go_collect_buttons) == 0:
        raise Exception("Visible Go Collect Money button was not found.")

    go_collect_button = visible_go_collect_buttons[0]
    print(f"Go Collect Money button text: '{go_collect_button.inner_text().strip()}'")

    go_collect_button.scroll_into_view_if_needed()
    print("Clicking Go collect money...")

    go_collect_button.click()
    print("Go collect money button clicked successfully.")

    print()
    print("Waiting for payment screen...")

    try:
        context = page.context

        print(f"Playwright currently sees {len(context.pages)} page(s).")

        active_pages = [p for p in context.pages if not p.is_closed()]

        print(f"Active Playwright pages found: {len(active_pages)}")

        if len(active_pages) == 0:
            raise Exception(
                "No active Playwright page was found after clicking Go collect money."
            )

        # Use the last active page — CRM may have opened/replaced the payment page.
        page = active_pages[-1]

        print("Active payment page selected.")
        print(f"Payment page URL: {page.url}")

    except Exception as e:
        print(f"Could not identify the active payment page: {e}")
        raise

    print()
    print("Selecting payment method: Bank")
    print()
    print("Waiting for Payment Methods control to appear...")

    payment_selects = page.locator("div.el-select.avue-select")

    for attempt in range(30):
        select_count = payment_selects.count()

        print(f"Payment UI check {attempt + 1}/30: Avue selects={select_count}")

        if select_count >= 4:
            break

        page.wait_for_timeout(1000)

    else:
        raise Exception(
            "Payment Methods control did not appear within 30 seconds."
        )

    print()
    print("Payment Methods control is now present.")

    # Use the 4th Avue select as the Payment Methods field.
    payment_select = payment_selects.nth(3)

    print()
    print("Using the 4th Avue select as Payment Methods.")
    print(f"Payment select visible: {payment_select.is_visible()}")

    if not payment_select.is_visible():
        raise Exception("Payment Methods select is present but not visible.")

    print()
    print("Opening Payment Methods dropdown...")

    payment_select.click()
    print("Payment Methods dropdown opened.")

    page.wait_for_timeout(500)

    print()
    print("Selecting payment method: Bank")

    bank_option = page.locator('[role="option"]').filter(has_text="bank")

    print(f"Bank option elements found: {bank_option.count()}")

    if bank_option.count() == 0:
        print()
        print("Available payment options:")

        options = page.locator('[role="option"]')

        for i in range(options.count()):
            try:
                print(f"  - {options.nth(i).inner_text().strip()}")
            except Exception:
                pass

        raise Exception("Bank payment option could not be found.")

    bank_option.first.click()
    print("Bank payment method selected successfully.")

    page.wait_for_timeout(500)

    print()
    print("Looking for payment Submit button...")

    payment_submit_button = page.locator(
        "button.el-button.el-button--primary"
    ).filter(has_text="submit")

    print(f"Payment Submit buttons found: {payment_submit_button.count()}")

    visible_payment_submit_buttons = []

    for i in range(payment_submit_button.count()):
        try:
            button = payment_submit_button.nth(i)
            if button.is_visible():
                visible_payment_submit_buttons.append(button)
        except Exception:
            continue

    print(
        f"Visible payment Submit buttons found: {len(visible_payment_submit_buttons)}"
    )

    if len(visible_payment_submit_buttons) == 0:
        raise Exception("Visible payment Submit button was not found.")

    payment_submit_button = visible_payment_submit_buttons[-1]
    print(f"Payment Submit button text: '{payment_submit_button.inner_text().strip()}'")

    payment_submit_button.scroll_into_view_if_needed()
    print()
    print("Clicking payment Submit...")

    payment_submit_button.click()
    print("Payment Submit clicked successfully.")

    print()
    print("Waiting for payment response...")

    try:
        page.wait_for_load_state("domcontentloaded", timeout=5000)
    except Exception:
        pass

    page.wait_for_timeout(2000)

    print()
    print("=" * 60)
    print("PAYMENT TEST COMPLETED")
    print("=" * 60)

    try:
        body_text = page.locator("body").inner_text()
        print()
        print("Visible page text after payment:")
        print(body_text[-4000:])
    except Exception as e:
        print(f"Could not read page text after payment: {e}")


# ============================================================
# EXCEL -> CRM ORDER PROCESSING HELPERS
# ============================================================



def get_crm_search_text(line):
    """
    Determine what to enter into the CRM 'The title of the product' search field.
    """
    description = str(line["description"]).strip()
    item_no = str(line["item_no"]).strip().upper()

    if not description:
        raise Exception(
            f"Cannot determine CRM search text for item {item_no} "
            f"because the Excel description is blank."
        )

    description_upper = description.upper()

    # --------------------------------------------------------
    # FRENCIA BASIN FAUCETS (BF-)
    # BF-22S in Excel -> BF-225 in CRM
    # --------------------------------------------------------

    if "BF-22S" in description_upper:
        return "BF-225"

    bf_match = re.search(
        r"\bBF-(\d{3})(?:[A-Z])?\b",
        f"{item_no} {description_upper}"
    )

    if bf_match:
        return f"BF-{bf_match.group(1)}"

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
    # OTHER FRENCIA PRODUCTS
    # --------------------------------------------------------

    if "FRENCIA" in description_upper:
        embedded_code = extract_embedded_crm_code(description)

        if embedded_code:
            if embedded_code in {
                "WC-006T",
                "WC-006P",
                "WC-008T",
                "WC-008P",
                "WC-009T",
                "WC-009P",
                "WC-664T",
                "WC-664P",
                "WC-664W",
            }:
                return embedded_code[:-1]

            return embedded_code

        if item_no:
            return item_no

        raise Exception(
            f"FRENCIA product does not contain a recognizable CRM code:\n{description}"
        )

    # --------------------------------------------------------
    # FT TILES
    # Skip the first 5 characters, then take the first 5 digits.
    # e.g. "FT FGE24101J WH ..." -> skip "FT FG" -> "24101"
    #
    # EXCEPTION — MRP / MR items:
    # Item codes starting with MRP or MR use the 6 digits
    # immediately after the MRP/MR prefix in the item number.
    # e.g. item "MRP612002Y" -> "612002"
    #      item "MR612009Y"  -> "612009"
    # --------------------------------------------------------

    if description_upper.startswith("FT "):
        # MRP prefix — 6 digits after "MRP"
        if item_no.startswith("MRP"):
            mrp_digits = re.search(r"\d{6}", item_no[3:])
            if mrp_digits:
                return mrp_digits.group(0)

        # MR prefix (no P) — 6 digits after "MR"
        if item_no.startswith("MR"):
            mr_digits = re.search(r"\d{6}", item_no[2:])
            if mr_digits:
                return mr_digits.group(0)

        # Standard rule: skip first 5 chars, take first 5 digits.
        digits = re.search(r"\d{5}", description[5:])
        if digits:
            return digits.group(0)

        # Fallback: first 5-digit number anywhere.
        fallback = re.search(r"\d{5}", description)
        if fallback:
            return fallback.group(0)

    # --------------------------------------------------------
    # WT TILES
    # Skip the first 6 characters, then take the first 5 digits.
    # e.g. "WT ABCD 12345 ..." -> skip "WT ABC" -> "12345"
    # --------------------------------------------------------

    if description_upper.startswith("WT "):
        digits = re.search(r"\d{5}", description[6:])

        if digits:
            return digits.group(0)

        # Fallback: first 5-digit number anywhere.
        fallback = re.search(r"\d{5}", description)
        if fallback:
            return fallback.group(0)

    # --------------------------------------------------------
    # NORMAL PRODUCTS
    # --------------------------------------------------------

    return description


def get_crm_price(line, order_lines=None):
    """
    Determine the CRM unit price for a normal product.
    Special-product exceptions are handled inline.
    """
    item_no = str(line["item_no"]).strip()
    description = str(line["description"]).upper()
    unit_price_excel = float(line["unit_price_excel"])
    amount_excel = line.get("amount_excel")

    if not unit_price_excel.is_integer() and amount_excel is not None:
        excel_price = float(amount_excel)
    else:
        excel_price = unit_price_excel

    # Frencia and Asian toilet products use the Excel/BC unit price.
    if (
        "FRENCIA" in description
        or "ASIAN TOILET" in description
        or "SQUATTING PAN" in description
        or "C.T. PAN" in description
    ):
        return excel_price

    # SC-001 fixed price.
    if item_no == "SC-001":
        return 1000

    order_item_nos = {
        str(order_line["item_no"]).strip().upper()
        for order_line in (order_lines or [])
    }

    # WC-006T + WC-006P + SC-001 / WC-008T + WC-008P + SC-001 / WC-009T + WC-009P + SC-001
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
    }

    if item_no in matching_pairs:
        matching_item = matching_pairs[item_no]

        if matching_item in order_item_nos:
            return excel_price / 2

    # Standard digit-prefix pricing.
    # For FT/WT items whose embedded code starts with a known prefix,
    # use the Excel Amount column value directly instead of a fixed price.
    if "FT " in description or "WT " in description:
        embedded_code = re.search(
            r"(?<!\d)(?:24|33|44|66|61|40|55|36|45)\d{3,4}(?!\d)",
            description
        )

        if embedded_code:
            # Use the Unit Price Incl. VAT value for these tile codes.
            return unit_price_excel

    return excel_price


def parse_inventory_value(inventory_text):
    """Convert the CRM Dealer warehouse inventory text into a numeric quantity."""
    text = str(inventory_text).strip()

    if not text:
        return 0

    cleaned = text.replace(",", "").strip()

    try:
        return float(cleaned)
    except Exception as e:
        raise Exception(
            f"Could not parse CRM inventory value '{inventory_text}'."
        ) from e


def print_excel_order_plan(order):
    """Display the exact order data that will be sent to the CRM."""
    print()
    print("=" * 80)
    print(f"ORDER PLAN: {order['document_no']}")
    print("=" * 80)
    print(f"Sell-to Customer No.: {order['customer_no']}")
    print(f"Lines: {len(order['lines'])}")

    for index, line in enumerate(order["lines"], start=1):
        search_text = get_crm_search_text(line)
        crm_price = get_crm_price(line, order["lines"])

        print()
        print(f"LINE {index}")
        print(f"  BC Item: {line['item_no']}")
        print(f"  Description: {line['description']}")
        print(f"  CRM Search: {search_text}")
        print(f"  Quantity: {line['quantity']}")
        print(f"  Excel Price: {line['unit_price_excel']}")
        print(f"  CRM Price: {crm_price}")
        print(f"  UOM: {line['unit_of_measure']}")

    print("=" * 80)


def main():
    if EXCEL_PREVIEW_ONLY:
        orders = load_excel_orders(EXCEL_FILE)

        print()
        print("=" * 70)
        print("EXCEL ORDER PREVIEW")
        print("=" * 70)

        preview_count = min(len(orders), 10)

        for index, order in enumerate(
            list(orders.values())[:preview_count], start=1
        ):
            print()
            print(f"ORDER {index}:")
            print(f"Document No.: {order['document_no']}")
            print(f"Customer No.: {order['customer_no']}")
            print(f"Lines: {len(order['lines'])}")

            for line in order["lines"]:
                print(
                    f"  {line['item_no']} | {line['description']} | "
                    f"Qty={line['quantity']} | Excel Price={line['unit_price_excel']}"
                )

        print()
        print("EXCEL PREVIEW COMPLETED.")

        return

    with sync_playwright() as p:
        browser = None
        context = None

        try:
            print("Opening CRM...")

            browser = p.chromium.launch(
                executable_path=BRAVE_PATH,
                headless=False,
                args=["--start-maximized"],
            )

            context = browser.new_context(no_viewport=True)
            page = context.new_page()

            page.goto(CRM_URL, wait_until="domcontentloaded", timeout=60000)

            print()
            print("Please log in manually and complete MFA if required.")

            # ------------------------------------------------
            # MANUAL LOGIN
            # ------------------------------------------------

            page.get_by_text(
                "Customer/Sales Management", exact=True
            ).wait_for(state="visible", timeout=300000)

            print("LOGIN CONFIRMED")

            # ------------------------------------------------
            # CUSTOMER / SALES MANAGEMENT
            # ------------------------------------------------

            page.get_by_text(
                "Customer/Sales Management", exact=True
            ).click(force=True)

            page.wait_for_timeout(1000)
            print("Customer/Sales Management opened.")

            # ------------------------------------------------
            # CUSTOMER ORDER MANAGEMENT
            # ------------------------------------------------

            customer_order_management = page.locator(
                'li.el-menu-item[data-track-args-code="custOrderSearch"]'
            )

            customer_order_management.wait_for(state="visible", timeout=15000)
            customer_order_management.click(force=True)

            page.wait_for_timeout(1200)
            print("Customer Order management opened.")

            # ------------------------------------------------
            # CREATE ORDER
            # ------------------------------------------------

            print()
            print("=" * 60)
            print("OPENING CREATE AN ORDER")
            print("=" * 60)

            create_order = page.get_by_role(
                "button", name="Create an order", exact=True
            )

            create_order.wait_for(state="visible", timeout=15000)
            print("Create an order button found.")

            pages_before = context.pages

            print(f"Browser pages before clicking: {len(pages_before)}")

            current_page_before = page

            print("Clicking Create an order...")
            create_order.click(force=True)
            print("Create an order clicked.")

            page.wait_for_timeout(1500)

            pages_after = context.pages

            print(f"Browser pages after clicking: {len(pages_after)}")

            new_pages = [p for p in pages_after if p not in pages_before]

            if new_pages:
                print()
                print("NEW ORDER TAB DETECTED.")

                order_page = new_pages[-1]

                try:
                    order_page.wait_for_load_state(
                        "domcontentloaded", timeout=30000
                    )
                except PlaywrightTimeoutError:
                    print(
                        "Order tab did not finish domcontentloaded within timeout; "
                        "continuing because the page exists."
                    )
            else:
                print()
                print("No new tab detected.")
                print(
                    "Checking whether the existing page became the order page..."
                )

                order_page = current_page_before
                order_page.wait_for_timeout(1500)

            print()
            print(f"Current order page URL: {order_page.url}")

            order_page.on("close", lambda: print("!!! CRM ORDER PAGE CLOSED !!!"))
            order_page.on("crash", lambda: print("!!! CRM ORDER PAGE CRASHED !!!"))

            # Verify the order-entry page before continuing.
            order_page_indicator = order_page.locator(
                'input[placeholder="Please input Sell customers"]'
            )

            try:
                order_page_indicator.wait_for(state="visible", timeout=15000)
                print()
                print("ORDER ENTRY PAGE CONFIRMED.")
            except PlaywrightTimeoutError:
                print()
                print("Sell customers field was not immediately found.")

                order_page.wait_for_timeout(2000)

                try:
                    order_page_indicator.wait_for(state="visible", timeout=10000)
                    print("ORDER ENTRY PAGE CONFIRMED AFTER ADDITIONAL WAIT.")
                except PlaywrightTimeoutError:
                    raise Exception(
                        "Create an order was clicked, but the CRM order-entry page "
                        f"could not be confirmed. Current URL: {order_page.url}"
                    )

            print()
            print(f"Order page: {order_page.url}")

            # ------------------------------------------------
            # ORDER TYPE
            # ------------------------------------------------

            print()
            print("Selecting Retail order...")

            order_type_item = order_page.locator(".el-form-item").filter(
                has_text="Order Type:"
            )

            order_type_select = order_type_item.locator(".el-select").first
            order_type_select.click(force=True)

            order_page.wait_for_timeout(500)

            retail_option = order_page.get_by_text("Retail order", exact=True)
            retail_visible = None

            for i in range(retail_option.count()):
                candidate = retail_option.nth(i)
                try:
                    if candidate.is_visible():
                        retail_visible = candidate
                        break
                except Exception:
                    pass

            if retail_visible is None:
                raise Exception("Retail order option not found.")

            retail_visible.click(force=True)
            order_page.wait_for_timeout(500)

            print("Order Type selected: Retail order")

            # ------------------------------------------------
            # SELL CUSTOMER
            # ------------------------------------------------

            print()
            print("Entering Sell customers name...")

            customer_name = get_next_customer_name()
            print(f"Entering Sell customers name: {customer_name}")

            # Create a DB run record as soon as we have the customer name.
            _db_run_id = _db.create_run(
                excel_filename=str(EXCEL_FILE.name) if EXCEL_FILE else "",
                customer_name=customer_name,
            )

            sell_customer = order_page.locator(
                'input[placeholder="Please input Sell customers"]'
            )

            sell_customer.wait_for(state="visible", timeout=15000)
            sell_customer.fill(customer_name)

            print(f"Sell customers entered: {customer_name}")

            # ------------------------------------------------
            # AUTOMATIC OUTBOUND SHIPMENT
            # ------------------------------------------------

            print()
            print("Checking Automatic outbound shipment...")

            shipment_label = order_page.locator("label.el-checkbox").filter(
                has_text="Automatic outbound shipment"
            )

            shipment_label.wait_for(state="visible", timeout=15000)

            shipment_checkbox = shipment_label.locator(
                'input.el-checkbox__original[type="checkbox"]'
            )

            visible_checkbox = shipment_label.locator(".el-checkbox__inner")

            initial_state = shipment_checkbox.is_checked()
            print(f"Initial checkbox state: {initial_state}")

            if not initial_state:
                visible_checkbox.scroll_into_view_if_needed()
                print("Clicking Automatic outbound shipment...")
                visible_checkbox.click(force=True)
                print("Automatic outbound shipment clicked.")
            else:
                print("Automatic outbound shipment was already selected.")

            # ------------------------------------------------
            # HANDLE POSSIBLE CONFIRMATION DIALOG
            # ------------------------------------------------

            print()
            print("Checking for automatic shipment confirmation dialog...")

            shipment_confirmed = False

            order_page.wait_for_timeout(700)

            try:
                confirm_buttons = order_page.get_by_role(
                    "button", name="Confirm", exact=True
                )

                visible_confirm_buttons = []

                for i in range(confirm_buttons.count()):
                    button = confirm_buttons.nth(i)
                    try:
                        if button.is_visible():
                            visible_confirm_buttons.append(button)
                    except Exception:
                        continue

                print(f"Visible Confirm buttons found: {len(visible_confirm_buttons)}")

                if visible_confirm_buttons:
                    print("Shipment confirmation dialog detected.")
                    visible_confirm_buttons[-1].click(force=True)
                    shipment_confirmed = True
                    print("Automatic outbound shipment confirmed.")

            except Exception as e:
                print(f"Confirmation dialog check produced: {e}")

            if not shipment_confirmed:
                print("No shipment confirmation dialog detected.")

            # ------------------------------------------------
            # FINAL VERIFICATION
            # ------------------------------------------------

            print()
            print("Verifying Automatic outbound shipment state...")

            order_page.wait_for_timeout(1000)

            final_state = shipment_checkbox.is_checked()
            print(f"Checkbox state after confirmation: {final_state}")

            if not final_state:
                raise Exception(
                    "Automatic outbound shipment checkbox is not selected after confirmation."
                )

            print()
            print("Automatic outbound shipment step complete.")

            # ------------------------------------------------
            # LOAD REAL EXCEL ORDER DATA
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("LOADING REAL EXCEL ORDER DATA")
            print("=" * 70)

            excel_orders = load_excel_orders(EXCEL_FILE)

            if not excel_orders:
                raise Exception("No usable Excel order data was found.")

            # Persist Excel order lines to the database.
            _db.save_excel_order_lines(_db_run_id, excel_orders)

            # Flatten all Excel order lines into one CRM batch.
            excel_lines = []

            for document_no, order in excel_orders.items():
                lines = order["lines"]

                print()
                print(f"Excel Document No. {document_no}: {len(lines)} line(s)")

                for line in lines:
                    excel_lines.append(line)

            print()
            print(f"Total Excel orders loaded: {len(excel_orders)}")
            print(f"Total Excel item lines loaded: {len(excel_lines)}")

            if not excel_lines:
                raise Exception(
                    "Excel file was loaded, but no item lines are available for CRM entry."
                )

            # ------------------------------------------------
            # ADD ALL EXCEL PRODUCTS TO CRM
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("ADDING EXCEL PRODUCTS TO CRM")
            print("=" * 70)

            for index, line in enumerate(excel_lines, start=1):
                document_no = str(line["document_no"]).strip()
                item_no = str(line["item_no"]).strip()
                description = str(line["description"]).strip()
                quantity = line["quantity"]
                search_text = get_crm_search_text(line)

                print()
                print("-" * 70)
                print(f"EXCEL LINE {index} / {len(excel_lines)}")
                print(f"Document No.: {document_no}")
                print(f"Item No.: {item_no}")
                print(f"Description: {description}")
                print(f"Quantity: {quantity}")
                print(f"CRM Search: {search_text}")

                # Snapshot order lines before product selection.
                before_order_lines = get_order_line_records(order_page)

                print()
                print(
                    f"Existing CRM order lines before search: {len(before_order_lines)}"
                )

                open_product_window(order_page)
                search_product(order_page, search_text)

                selected_products = select_all_product_results(
                    order_page, search_text
                )

                print()
                print(f"Selected CRM search results: {len(selected_products)}")

                order_page = confirm_product_selection(order_page)

                order_page.wait_for_timeout(1000)

                after_order_lines = get_order_line_records(order_page)

                print()
                print(f"CRM order lines after Confirm: {len(after_order_lines)}")

                new_order_lines = get_new_order_lines(
                    before_order_lines, after_order_lines
                )

                print()
                print(
                    f"New CRM order lines created for this BC item: {len(new_order_lines)}"
                )

                if not new_order_lines:
                    raise Exception(
                        "CRM Confirm completed, but no newly created order lines were detected."
                    )

                # ------------------------------------------------
                # DELETE PENDING ZERO-INVENTORY LINES NOW THAT THE
                # NEXT NORMAL CRM ITEM HAS BEEN ADDED
                # ------------------------------------------------

                if PENDING_ZERO_INVENTORY_DELETIONS:
                    pending_zero_deletions = PENDING_ZERO_INVENTORY_DELETIONS.copy()
                    PENDING_ZERO_INVENTORY_DELETIONS.clear()

                    # Build a set of (item_no, product_title) keys for the
                    # lines just added for the current Excel item.
                    # Any pending deletion whose key matches a newly added line
                    # must be skipped — it would delete the fresh line we just
                    # added, not a leftover from the previous item.
                    new_line_keys = {
                        (
                            str(nl["item_no"]).strip().upper(),
                            str(nl["product_title"]).strip().upper(),
                        )
                        for nl in new_order_lines
                    }

                    # Highest occurrence first prevents duplicate product
                    # identities from shifting before deletion.
                    pending_zero_deletions.sort(
                        key=lambda d: (
                            str(d.get("item_no", "")),
                            str(d.get("product_title", "")),
                            int(d.get("occurrence", 0)),
                        ),
                        reverse=True,
                    )

                    for pending_line in pending_zero_deletions:
                        pending_item = str(pending_line.get("item_no", "")).strip()
                        pending_product = str(
                            pending_line.get("product_title", "")
                        ).strip()
                        pending_occurrence = int(pending_line.get("occurrence", 0))

                        pending_key = (
                            pending_item.upper(),
                            pending_product.upper(),
                        )

                        # Skip if this product was just added as a new line
                        # for the current Excel item.
                        if pending_key in new_line_keys:
                            print()
                            print(
                                f"SKIPPING pending deletion — product was "
                                f"re-added for the current item:\n"
                                f"{pending_item} | {pending_product}"
                            )
                            continue

                        print()
                        print("=" * 70)
                        print(
                            "DELETING PENDING ZERO-INVENTORY LINE "
                            "AFTER NEXT ITEM WAS ADDED"
                        )
                        print("=" * 70)
                        print(f"Item: {pending_item}")
                        print(f"Product: {pending_product}")
                        print(f"Original occurrence: {pending_occurrence}")

                        deleted = delete_order_line_occurrence(
                            order_page,
                            pending_item,
                            pending_product,
                            pending_occurrence,
                        )

                        if not deleted:
                            print()
                            print(
                                f"WARNING: Pending zero-inventory CRM line "
                                f"could not be deleted — skipping and continuing.\n"
                                f"{pending_item} | {pending_product} | "
                                f"Occurrence={pending_occurrence}"
                            )
                            print(
                                "The safety-net cleanup at the end of the "
                                "run will handle any remaining zero-inventory lines."
                            )
                        else:
                            print()
                            print(
                                f"Pending zero-inventory line deleted successfully: "
                                f"{pending_item} | {pending_product}"
                            )

                # ------------------------------------------------
                # ALLOCATE BC QUANTITY ACROSS CRM LINES
                # ------------------------------------------------

                # Re-snapshot order lines after any pending deletions.
                # Pending deletions can remove earlier occurrences of the
                # same product, shifting occurrence numbers of newly added
                # lines. A fresh snapshot ensures correct occurrence indices
                # are passed to allocate_quantity_across_new_lines.
                after_order_lines = get_order_line_records(order_page)
                new_order_lines = get_new_order_lines(
                    before_order_lines, after_order_lines
                )
                print()
                print(
                    f"New CRM order lines (after pending deletions): "
                    f"{len(new_order_lines)}"
                )

                current_order_lines = excel_orders[document_no]["lines"]

                zero_report_start = len(ZERO_INVENTORY_REPORT_LINES)

                paired_model = str(search_text).strip().upper()

                paired_toilet_models = {
                    "WC-004", "WC-006", "WC-008", "WC-009",
                    "WC-029", "WC-100", "WC-5096", "WC-664",
                }

                paired_models_with_sc = {"WC-006", "WC-008", "WC-009"}

                if paired_model == "SQ":
                    # ------------------------------------------------
                    # SPECIAL ASIAN TOILET PROCESS
                    # ------------------------------------------------

                    print()
                    print("=" * 70)
                    print("SPECIAL ASIAN TOILET PROCESS")
                    print("=" * 70)

                    asian_p_lines = [
                        r for r in new_order_lines
                        if "SQ-001P" in str(r["product_title"]).upper()
                    ]

                    asian_t_lines = [
                        r for r in new_order_lines
                        if "SQ-001T" in str(r["product_title"]).upper()
                    ]

                    print(f"SQ-001P CRM lines found: {len(asian_p_lines)}")
                    print(f"SQ-001T CRM lines found: {len(asian_t_lines)}")

                    if not asian_p_lines:
                        raise Exception("No SQ-001P CRM lines were created.")

                    if not asian_t_lines:
                        raise Exception("No SQ-001T CRM lines were created.")

                    asian_base_price = get_crm_price(line, current_order_lines)
                    asian_component_price = asian_base_price / 2

                    print()
                    print(f"Asian Toilet Excel price: {asian_base_price}")
                    print(f"SQ-001P price: {asian_component_price}")
                    print(f"SQ-001T price: {asian_component_price}")

                    # Delete non-pair SQ search results.
                    asian_pair_lines = asian_p_lines + asian_t_lines

                    asian_extra_lines = [
                        r for r in new_order_lines if r not in asian_pair_lines
                    ]

                    asian_extra_lines.sort(
                        key=lambda r: (r["key"], r["occurrence"]), reverse=True
                    )

                    print()
                    print("REMOVING NON-PAIR ASIAN TOILET LINES")

                    for line_record in asian_extra_lines:
                        delete_order_line_occurrence(
                            order_page,
                            line_record["item_no"],
                            line_record["product_title"],
                            line_record["occurrence"],
                        )

                    retained_lines = []

                    retained_asian_p_lines = allocate_quantity_across_new_lines(
                        order_page, asian_p_lines, quantity, asian_component_price
                    )
                    retained_lines.extend(retained_asian_p_lines)

                    retained_asian_t_lines = allocate_quantity_across_new_lines(
                        order_page, asian_t_lines, quantity, asian_component_price
                    )
                    retained_lines.extend(retained_asian_t_lines)

                elif paired_model in paired_toilet_models:
                    # ------------------------------------------------
                    # SPECIAL PAIRED TOILET PROCESS
                    # ------------------------------------------------

                    print()
                    print("=" * 70)
                    print("SPECIAL PAIRED TOILET PROCESS")
                    print("=" * 70)

                    paired_p_lines = [
                        r for r in new_order_lines
                        if f"{paired_model}P" in str(r["product_title"]).upper()
                    ]

                    paired_t_lines = [
                        r for r in new_order_lines
                        if f"{paired_model}T" in str(r["product_title"]).upper()
                    ]

                    print(f"{paired_model}P CRM lines found: {len(paired_p_lines)}")
                    print(f"{paired_model}T CRM lines found: {len(paired_t_lines)}")

                    if not paired_p_lines:
                        raise Exception(
                            f"No {paired_model}P CRM lines were created."
                        )

                    if not paired_t_lines:
                        raise Exception(
                            f"No {paired_model}T CRM lines were created."
                        )

                    if paired_model in paired_models_with_sc:
                        print()
                        print("=" * 70)
                        print("SPECIAL SC-001 COMPANION PROCESS")
                        print("=" * 70)

                        print()
                        print("Adding required SC-001 companion...")

                        before_sc_lines = get_order_line_records(order_page)

                        open_product_window(order_page)
                        search_product(order_page, "SC-001")

                        try:
                            selected_sc_products = select_all_product_results(
                                order_page, "SC-001"
                            )

                            print(
                                f"Selected SC-001 CRM results: {len(selected_sc_products)}"
                            )

                            order_page = confirm_product_selection(order_page)
                            order_page.wait_for_timeout(1000)

                            after_sc_lines = get_order_line_records(order_page)

                            sc_new_lines = get_new_order_lines(
                                before_sc_lines, after_sc_lines
                            )

                            print(f"New SC-001 CRM lines created: {len(sc_new_lines)}")

                            if not sc_new_lines:
                                print(
                                    "WARNING: SC-001 confirmed but no new line "
                                    "detected — skipping SC-001 companion."
                                )
                                sc_new_lines = []

                        except Exception as sc_error:
                            print()
                            print(
                                f"WARNING: SC-001 companion could not be added: "
                                f"{sc_error}"
                            )
                            print(
                                "Continuing without SC-001 — P and T components "
                                "will still be allocated."
                            )
                            sc_new_lines = []

                    # Calculate paired component price.
                    unit_price_excel = float(line["unit_price_excel"])
                    amount_excel = line.get("amount_excel")

                    if not unit_price_excel.is_integer() and amount_excel is not None:
                        paired_base_price = float(amount_excel)
                    else:
                        paired_base_price = unit_price_excel

                    no_sc_models = {"WC-004", "WC-029", "WC-100", "WC-5096", "WC-664"}

                    if paired_model in no_sc_models:
                        paired_component_price = paired_base_price / 2
                    else:
                        paired_component_price = (paired_base_price - 1000) / 2

                    if paired_component_price < 0:
                        raise Exception(
                            f"Invalid paired-product price for {paired_model}: "
                            f"{paired_component_price}"
                        )

                    print()
                    print(f"Paired base price: {paired_base_price}")
                    print(f"{paired_model}P price: {paired_component_price}")
                    print(f"{paired_model}T price: {paired_component_price}")

                    if paired_model not in no_sc_models:
                        print("SC-001 price: 1000")

                    retained_lines = []

                    retained_p_lines = allocate_quantity_across_new_lines(
                        order_page, paired_p_lines, quantity, paired_component_price
                    )
                    retained_lines.extend(retained_p_lines)

                    retained_t_lines = allocate_quantity_across_new_lines(
                        order_page, paired_t_lines, quantity, paired_component_price
                    )
                    retained_lines.extend(retained_t_lines)

                    if paired_model in paired_models_with_sc and sc_new_lines:
                        retained_sc_lines = allocate_quantity_across_new_lines(
                            order_page, sc_new_lines, quantity, 1000
                        )
                        retained_lines.extend(retained_sc_lines)

                else:
                    # ------------------------------------------------
                    # NORMAL PRODUCT ALLOCATION
                    # ------------------------------------------------

                    crm_price = get_crm_price(line, current_order_lines)

                    retained_lines = allocate_quantity_across_new_lines(
                        order_page, new_order_lines, quantity, crm_price
                    )

                # ------------------------------------------------
                # COLLECT POSTING REPORT ROWS
                # ------------------------------------------------

                customer_no = str(
                    excel_orders[document_no].get("customer_no", "")
                ).strip()

                for retained_line in retained_lines:
                    POSTING_REPORT_ROWS.append({
                        "posting_date": time.strftime("%Y-%m-%d"),
                        "customer_no": customer_no,
                        "document_no": document_no,
                        "excel_item_no": item_no,
                        "excel_description": description,
                        "excel_qty": float(quantity),
                        "excel_unit_price": float(line["unit_price_excel"]),
                        "crm_item_no": str(retained_line["item_no"]).strip(),
                        "crm_product": str(retained_line["product_title"]).strip(),
                        "crm_inventory": float(retained_line["inventory"]),
                        "allocated_qty": float(retained_line["allocated_quantity"]),
                        "crm_unit_price": float(retained_line["unit_price"]),
                        "inventory_status": (
                            "ZERO INVENTORY"
                            if float(retained_line["inventory"]) <= 0
                            else (
                                "INSUFFICIENT INVENTORY"
                                if float(retained_line["allocated_quantity"])
                                < float(quantity)
                                else "ENOUGH INVENTORY"
                            )
                        ),
                    })

                for zero_line in ZERO_INVENTORY_REPORT_LINES[zero_report_start:]:
                    POSTING_REPORT_ROWS.append({
                        "posting_date": time.strftime("%Y-%m-%d"),
                        "customer_no": customer_no,
                        "document_no": document_no,
                        "excel_item_no": item_no,
                        "excel_description": description,
                        "excel_qty": float(quantity),
                        "excel_unit_price": float(line["unit_price_excel"]),
                        "crm_item_no": str(zero_line["item_no"]).strip(),
                        "crm_product": str(zero_line["product_title"]).strip(),
                        "crm_inventory": float(zero_line["inventory"]),
                        "allocated_qty": 0.0,
                        "crm_unit_price": float(zero_line["unit_price"]),
                        "inventory_status": "ZERO INVENTORY",
                    })

                print()
                print(f"Retained CRM lines for BC item: {len(retained_lines)}")
                print("Product inventory allocation completed.")

            # ------------------------------------------------
            # SHOW CRM ORDER LINES
            # ------------------------------------------------

            remove_zero_inventory_order_lines(order_page)

            print()
            print("=" * 70)
            print("CRM ORDER LINES AFTER PRODUCT ENTRY")
            print("=" * 70)

            print_current_order_lines(order_page)

            # ------------------------------------------------
            # FINAL CRM ORDER INSPECTION
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("FINAL CRM ORDER STATE")
            print("=" * 70)

            print_current_order_lines(order_page)

            print()
            print("=" * 70)
            print("INSPECTING CALCULATED CRM VALUES")
            print("=" * 70)

            for line in excel_lines:
                description = str(line["description"]).strip()

                try:
                    inspect_order_row_values(order_page, description)
                except Exception as inspection_error:
                    print(f"Could not inspect {description}: {inspection_error}")

            # ------------------------------------------------
            # SUBMIT ORDER
            # ------------------------------------------------

            print()
            print("=" * 70)
            print("SUBMITTING EXCEL-DRIVEN CRM ORDER")
            print("=" * 70)

            test_submit_order(order_page)

            # ------------------------------------------------
            # CREATE POSTING REPORT AFTER SUCCESSFUL SUBMIT
            # ------------------------------------------------

            report_path = write_posting_report()

            # Save posting report rows to DB and mark run as successful.
            _db.save_posting_report_rows(_db_run_id, POSTING_REPORT_ROWS)
            _db.finish_run(
                run_id=_db_run_id,
                status="success",
                total_documents=len(excel_orders),
                total_lines=len(excel_lines),
                report_path=str(report_path) if report_path else None,
            )
            print(f"Run #{_db_run_id} saved to database.")

            print()
            print("=" * 70)
            print("EXCEL-DRIVEN CRM ORDER COMPLETED")
            print("=" * 70)
            print()
            print(f"Processed Excel documents: {len(excel_orders)}")
            print(f"Processed Excel item lines: {len(excel_lines)}")
            print()
            print("CRM calculations were left to the CRM.")
            print("The CRM page will remain open for inspection.")

            wait_for_enter("FINAL CRM PAGE")

        except Exception as e:
            print()
            print("=" * 70)
            print("TEST STOPPED BECAUSE OF AN ERROR")
            print("=" * 70)
            print(f"{type(e).__name__}: {e}")
            print()
            print("FULL TRACEBACK:")
            traceback.print_exc()
            print()
            print("The browser will remain open.")

            # Mark the run as failed in the database.
            try:
                _db.finish_run(
                    run_id=_db_run_id,
                    status="error",
                    error_message=f"{type(e).__name__}: {e}",
                )
            except Exception:
                pass

            wait_for_enter(
                "ERROR STATE - INSPECT THE CRM PAGE BEFORE CLOSING IT"
            )

            raise

        finally:
            # IMPORTANT: Do NOT close the browser automatically.
            # This is intentionally disabled while debugging the CRM workflow.
            print()
            print("Browser left open intentionally.")

            # No browser.close()
            # No context.close()


if __name__ == "__main__":
    main()
