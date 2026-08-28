import calendar
import io
import re
import zipfile
import xml.etree.ElementTree as XET
from copy import copy
from datetime import date, datetime
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.title("Monthly Line Listing Product Splitter")
st.caption("Upload one monthly line listing and download one workbook with a separate sheet for each Celix product.")

CELIX_PRODUCTS = [
    "abiraterone", "apixaban", "apremilast", "bexarotene", "brivaracetam", "clobazam",
    "clonazepam", "cyanocobalamin", "dabigatran", "dapagliflozin", "dimethyl fumarate",
    "edoxaban", "empagliflozin", "famotidine", "fesoterodine", "icatibant", "itraconazole",
    "linagliptin + metformin", "linagliptin", "metformin", "letrozole", "nintedanib",
    "pirfenidone", "raltegravir", "ranolazine", "rivaroxaban", "safinamide", "saxagliptin",
    "sitagliptin", "sacubitril + valsartan", "sacubritril + valsartan", "sacubitril",
    "sacubritril", "valsartan", "tamsulosin + solifenacin", "tapentadol", "ticagrelor",
    "tamsulosin", "solifenacin",
]

# Product-name variations that can appear in the line listing.
ALIASES = {
    "dabigatran": ["dabigatran etexilate"],
    "fesoterodine": ["fesoterodine fumarate"],
    "tamsulosin + solifenacin": ["tamsulosin/solifenacin", "solifenacin + tamsulosin", "vacit", "vecit"],
    "sacubitril + valsartan": ["sacubitril/valsartan", "sacubitril valsartan"],
    "sacubritril + valsartan": ["sacubritril/valsartan", "sacubritril valsartan"],
}


def norm(value):
    text = str(value or "").lower()
    text = re.sub(r"[^a-z0-9+]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def product_patterns():
    patterns = {}
    for product in CELIX_PRODUCTS:
        canonical = norm(product)
        patterns.setdefault(canonical, set()).add(canonical)
        for alias in ALIASES.get(product, []):
            patterns[canonical].add(norm(alias))
    # Combination products are checked first, preventing accidental single-ingredient matching.
    return {k: sorted(v, key=len, reverse=True) for k, v in patterns.items()}


PRODUCT_PATTERNS = product_patterns()


def exact_product_tokens(product_cell):
    # Product fields in the report are semicolon separated. Evaluate each product independently.
    return [norm(part) for part in re.split(r"[;\n]+", str(product_cell or "")) if norm(part)]


def matched_celix_products(product_cell):
    matches = []
    tokens = exact_product_tokens(product_cell)
    for canonical, aliases in sorted(PRODUCT_PATTERNS.items(), key=lambda item: len(item[0]), reverse=True):
        found = False
        for token in tokens:
            for alias in aliases:
                if token == alias or token.startswith(alias + " ") or re.search(rf"\b{re.escape(alias)}\b", token):
                    found = True
                    break
            if found:
                break
        if found and canonical not in matches:
            matches.append(canonical)
    # If a combination is present, do not also create separate sheets for its component words.
    if "tamsulosin + solifenacin" in matches:
        matches = [m for m in matches if m not in {"tamsulosin", "solifenacin"}]
    if "linagliptin + metformin" in matches:
        matches = [m for m in matches if m not in {"linagliptin", "metformin"}]
    if "sacubitril + valsartan" in matches or "sacubritril + valsartan" in matches:
        matches = [m for m in matches if m not in {"sacubitril", "sacubritril", "valsartan"}]
    return matches


def find_header_row(ws):
    for row in range(1, min(ws.max_row, 50) + 1):
        values = [norm(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 20) + 1)]
        if "safety report id" in values and "product name" in values:
            return row
    raise ValueError("Could not find the header row containing Safety Report ID and Product Name.")


def find_column(ws, header_row, name):
    wanted = norm(name)
    for col in range(1, ws.max_column + 1):
        if norm(ws.cell(header_row, col).value) == wanted:
            return col
    raise ValueError(f"Required column not found: {name}")


def clone_style(source, target):
    if source.has_style:
        target._style = copy(source._style)
    if source.hyperlink:
        target._hyperlink = copy(source.hyperlink)
    if source.comment:
        target.comment = copy(source.comment)


def safe_sheet_name(product, used):
    base = re.sub(r"[\\/*?:\[\]]", " ", product.upper()).strip()[:31] or "PRODUCT"
    name = base
    counter = 2
    while name.lower() in used:
        suffix = f"_{counter}"
        name = base[:31 - len(suffix)] + suffix
        counter += 1
    used.add(name.lower())
    return name


def product_display(product):
    return product.upper()


def set_top_details(ws, product, start_date, end_date):
    display = product_display(product)
    replacements = {
        3: f"Project Name: CELIXP-GB-({display})({display})",
        4: f"Active Ingredient: {display}",
        5: f"Proprietory Name: {display}",
        6: f"Drug Name: {display}",
        7: f"Period: ADR Receipt Date: :{start_date.strftime('%d-%b-%y')} To {end_date.strftime('%d-%b-%y')}",
    }
    for row, value in replacements.items():
        ws.cell(row, 1).value = value



def remove_blank_top_rows(ws):
    """Remove fully blank leading rows while preserving values, styles and merged ranges."""
    blank_count = 0
    for row in range(1, ws.max_row + 1):
        if all(ws.cell(row, col).value in (None, "") for col in range(1, ws.max_column + 1)):
            blank_count += 1
        else:
            break
    if blank_count == 0:
        return
    merges = [
        (rng.min_row, rng.min_col, rng.max_row, rng.max_col)
        for rng in ws.merged_cells.ranges
        if rng.max_row > blank_count
    ]
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    max_row, max_col = ws.max_row, ws.max_column
    for source_row in range(blank_count + 1, max_row + 1):
        target_row = source_row - blank_count
        for col in range(1, max_col + 1):
            source = ws.cell(source_row, col)
            target = ws.cell(target_row, col)
            target.value = source.value
            clone_style(source, target)
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    ws.delete_rows(max_row - blank_count + 1, blank_count)
    for min_row, min_col, max_row_merge, max_col_merge in merges:
        ws.merge_cells(
            start_row=max(1, min_row - blank_count),
            start_column=min_col,
            end_row=max_row_merge - blank_count,
            end_column=max_col_merge,
        )

def add_serial_number_column(ws, header_row):
    """Add Sl No only to the tabular area; leave the top report rows unchanged."""
    original_max_col = ws.max_column
    original_max_row = ws.max_row

    # Shift header and data cells one column to the right without inserting a full worksheet column.
    for row in range(original_max_row, header_row - 1, -1):
        for col in range(original_max_col, 0, -1):
            source = ws.cell(row, col)
            target = ws.cell(row, col + 1)
            target.value = source.value
            clone_style(source, target)
            source.value = None

    # Shift visible column widths for the table columns.
    for col in range(original_max_col, 0, -1):
        source_letter = get_column_letter(col)
        target_letter = get_column_letter(col + 1)
        ws.column_dimensions[target_letter].width = ws.column_dimensions[source_letter].width

    ws.cell(header_row, 1).value = "Sl No"
    clone_style(ws.cell(header_row, 2), ws.cell(header_row, 1))
    ws.column_dimensions["A"].width = 8


def normalize_xlsx_font_order(xlsx_bytes):
    """Preserve Excel-compatible font element order after openpyxl serialization."""
    source = io.BytesIO(xlsx_bytes)
    target = io.BytesIO()
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    order = {name: index for index, name in enumerate([
        "b", "i", "strike", "outline", "shadow", "condense", "extend",
        "sz", "color", "name", "family", "charset", "scheme", "u", "vertAlign"
    ])}
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                root = XET.fromstring(data)
                fonts = root.find(f"{{{ns}}}fonts")
                if fonts is not None:
                    for font in fonts.findall(f"{{{ns}}}font"):
                        children = list(font)
                        children.sort(key=lambda child: order.get(child.tag.rsplit("}", 1)[-1], 999))
                        font[:] = children
                data = XET.tostring(root, encoding="utf-8", xml_declaration=False)
            zout.writestr(item, data)
    return target.getvalue()

def build_split_workbook(uploaded_bytes, selected_year, selected_month):
    workbook = load_workbook(io.BytesIO(uploaded_bytes))
    source = workbook[workbook.sheetnames[0]]
    # The source report may contain a remote Print Date cell in the otherwise blank first row.
    # Remove that auxiliary row so the report title starts immediately at row 1.
    for col in range(1, source.max_column + 1):
        source.cell(1, col).value = None
    remove_blank_top_rows(source)
    header_row = find_header_row(source)
    product_col = find_column(source, header_row, "Product Name")
    date_col = find_column(source, header_row, "ADR Receipt Date/Time")

    first_day = date(selected_year, selected_month, 1)
    last_day = date(selected_year, selected_month, calendar.monthrange(selected_year, selected_month)[1])

    product_rows = {}
    for row in range(header_row + 1, source.max_row + 1):
        safety_id = source.cell(row, 1).value
        if safety_id in (None, ""):
            continue
        raw_date = source.cell(row, date_col).value
        row_date = None
        if isinstance(raw_date, datetime):
            row_date = raw_date.date()
        elif isinstance(raw_date, date):
            row_date = raw_date
        else:
            text = str(raw_date or "").strip()
            for fmt in ("%d-%b-%Y-%H:%M:%S", "%d-%b-%y-%H:%M:%S", "%d-%b-%Y", "%d-%b-%y"):
                try:
                    row_date = datetime.strptime(text, fmt).date()
                    break
                except ValueError:
                    pass
        if row_date and (row_date.year != selected_year or row_date.month != selected_month):
            continue
        for product in matched_celix_products(source.cell(row, product_col).value):
            product_rows.setdefault(product, []).append(row)

    template_name = source.title
    used_names = set()
    created = []
    for product in sorted(product_rows):
        ws = workbook.copy_worksheet(source)
        ws.title = safe_sheet_name(product, used_names)
        keep_rows = product_rows[product]
        # Remove existing data and rewrite matching rows from the original source sheet.
        if ws.max_row > header_row:
            ws.delete_rows(header_row + 1, ws.max_row - header_row)
        for output_index, source_row in enumerate(keep_rows, start=header_row + 1):
            for col in range(1, source.max_column + 1):
                src_cell = source.cell(source_row, col)
                dst_cell = ws.cell(output_index, col, src_cell.value)
                clone_style(src_cell, dst_cell)
            ws.row_dimensions[output_index].height = source.row_dimensions[source_row].height
            # Keep all report columns, but show only the current Celix product in Product Name.
            ws.cell(output_index, product_col).value = product_display(product)
        set_top_details(ws, product, first_day, last_day)
        add_serial_number_column(ws, header_row)
        for index, row in enumerate(range(header_row + 1, ws.max_row + 1), start=1):
            ws.cell(row, 1).value = index
            clone_style(ws.cell(row, 2), ws.cell(row, 1))
            ws.cell(row, 1).alignment = copy(ws.cell(row, 2).alignment)
        ws.freeze_panes = None
        ws.auto_filter.ref = None
        ws.sheet_view.showGridLines = source.sheet_view.showGridLines
        created.append((product_display(product), len(keep_rows)))

    workbook.remove(workbook[template_name])
    if not created:
        raise ValueError("No Celix product rows were found for the selected month.")
    output = io.BytesIO()
    workbook.save(output)
    result_bytes = normalize_xlsx_font_order(output.getvalue())
    return result_bytes, created


uploaded = st.file_uploader("Upload monthly line listing", type=["xlsx"])
current = date.today()
c1, c2 = st.columns(2)
with c1:
    selected_month = st.selectbox("Month", list(range(1, 13)), index=current.month - 1, format_func=lambda m: calendar.month_name[m])
with c2:
    selected_year = st.number_input("Year", min_value=2000, max_value=2100, value=current.year, step=1)

if uploaded and st.button("Generate product-wise workbook", type="primary"):
    try:
        result, summary = build_split_workbook(uploaded.getvalue(), int(selected_year), int(selected_month))
        st.success(f"Created {len(summary)} Celix product sheets.")
        st.dataframe({"Celix Product": [x[0] for x in summary], "Cases": [x[1] for x in summary]}, hide_index=True, use_container_width=True)
        filename = f"Celix_Product_Line_Listing_{int(selected_year)}_{int(selected_month):02d}.xlsx"
        st.download_button("Download segregated workbook", result, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as exc:
        st.error(f"Could not generate the workbook: {exc}")
