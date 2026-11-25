import io
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
import streamlit as st

# Paths to local templates
BASE_DIR = Path(__file__).parent
NZ_TEMPLATE_PATH = BASE_DIR / "templates" / "newzealand_template.xlsx"
CO_TEMPLATE_PATH = BASE_DIR / "templates" / "company_orders_template.xlsx"


def copy_cell_style(src_cell, dst_cell):
    """Copy font, fill, border, alignment, and number_format from src_cell to dst_cell."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def snapshot_style(cell) -> dict:
    """Capture style attributes of a cell so they remain valid after its workbook closes."""
    return {
        "font": copy(cell.font) if cell.has_style else None,
        "fill": copy(cell.fill) if cell.has_style else None,
        "border": copy(cell.border) if cell.has_style else None,
        "alignment": copy(cell.alignment) if cell.has_style else None,
        "number_format": cell.number_format if cell.has_style else "General",
    }


def apply_style(cell, style_dict: dict):
    """Apply a snapshotted style dictionary to a target cell."""
    if style_dict.get("font"):
        cell.font = copy(style_dict["font"])
    if style_dict.get("fill"):
        cell.fill = copy(style_dict["fill"])
    if style_dict.get("border"):
        cell.border = copy(style_dict["border"])
    if style_dict.get("alignment"):
        cell.alignment = copy(style_dict["alignment"])
    if style_dict.get("number_format"):
        cell.number_format = style_dict["number_format"]


def parse_invoice_excel(file_bytes: bytes, filename: str) -> dict:
    """
    Extract Q4 (Invoice No), S4 (Invoice Date), B24:W24 (Headers + styles),
    and B25:W25+ rows (values + row 25 styles) until Column B is empty.
    """
    wb_val = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_val = wb_val.active

    wb_sty = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws_sty = wb_sty.active

    invoice_no = ws_val["Q4"].value
    invoice_date = ws_val["S4"].value

    # Headers + styles from B24 to W24 (columns 2 to 23)
    headers = [ws_val.cell(24, col).value for col in range(2, 24)]
    header_styles = [snapshot_style(ws_sty.cell(24, col)) for col in range(2, 24)]
    data_styles = [snapshot_style(ws_sty.cell(25, col)) for col in range(2, 24)]
    header_height = ws_sty.row_dimensions[24].height
    data_height = ws_sty.row_dimensions[25].height

    # Data rows from row 25 onward until Column B (col 2) is empty
    rows = []
    r = 25
    while True:
        b_val = ws_val.cell(r, 2).value
        if b_val is None or str(b_val).strip() == "":
            break

        row_by_col = {}
        for col_idx in range(2, 24):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            row_by_col[col_letter] = ws_val.cell(r, col_idx).value

        rows.append(row_by_col)
        r += 1

    return {
        "filename": filename,
        "invoice_no": invoice_no,
        "invoice_date": invoice_date,
        "headers": headers,
        "header_styles": header_styles,
        "data_styles": data_styles,
        "header_height": header_height,
        "data_height": data_height,
        "rows": rows,
    }


def build_combined_raw_excel(parsed_files: list[dict]) -> bytes:
    """
    Output 1:
    - Header: ['Invoice No', 'Invoice Date'] + B24:W24 (with pink input header styling)
    - Data: B25:W25+ rows with Invoice No & Date repeated + input data cell borders/fonts
    - 1 blank row gap between different files
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Combined_Data"

    if not parsed_files:
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    first = parsed_files[0]
    header_styles = first["header_styles"]
    data_styles = first["data_styles"]

    # Write header once with input pink styling
    full_headers = ["Invoice No", "Invoice Date"] + first["headers"]
    ws.append(full_headers)
    if first["header_height"]:
        ws.row_dimensions[1].height = first["header_height"]

    # Apply B24's pink style to Invoice No & Date headers, and B24:W24 styles to C1:X1
    apply_style(ws.cell(1, 1), header_styles[0])
    apply_style(ws.cell(1, 2), header_styles[0])
    for offset, st_dict in enumerate(header_styles, start=3):
        apply_style(ws.cell(1, offset), st_dict)

    current_row = 2
    col_letters = [openpyxl.utils.get_column_letter(c) for c in range(2, 24)]

    for file_idx, item in enumerate(parsed_files):
        if file_idx > 0:
            # Leave a 1-row gap between files
            current_row += 1

        inv_no = item["invoice_no"]
        inv_date = item["invoice_date"]
        item_data_styles = item["data_styles"] or data_styles

        for row_dict in item["rows"]:
            if item["data_height"]:
                ws.row_dimensions[current_row].height = item["data_height"]

            c_inv = ws.cell(current_row, 1, inv_no)
            apply_style(c_inv, item_data_styles[1])  # clean thin-bordered text style

            c_date = ws.cell(current_row, 2, inv_date)
            apply_style(c_date, item_data_styles[1])
            if isinstance(inv_date, datetime):
                c_date.number_format = "dd-mm-yyyy"

            for idx, col_let in enumerate(col_letters):
                cell = ws.cell(current_row, idx + 3, row_dict.get(col_let))
                apply_style(cell, item_data_styles[idx])

            current_row += 1

    # Auto-fit column widths lightly
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 35)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_newzealand_excel(parsed_files: list[dict]) -> bytes:
    """
    Output 2:
    - Uses templates/newzealand_template.xlsx (keeps headers in A3:J3)
    - Maps: Q4, S4, I, C, L, M, F, S, T, U -> Columns A..J starting at row 4
    - Updates J2 =SUM(J4:J<last_row>)
    """
    wb = openpyxl.load_workbook(NZ_TEMPLATE_PATH)
    ws = wb.active

    template_row_styles = [snapshot_style(ws.cell(4, c)) for c in range(1, 11)]

    if ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)

    current_row = 4
    for item in parsed_files:
        inv_no = item["invoice_no"]
        inv_date = item["invoice_date"]

        for r in item["rows"]:
            mapped_values = [
                inv_no,        # A: Q4 (Header / Invoice No)
                inv_date,      # B: S4 (invoice date)
                r.get("I"),    # C: Customer PO NO
                r.get("C"),    # D: Design No
                r.get("L"),    # E: Barcode
                r.get("M"),    # F: Item Description
                r.get("F"),    # G: Gold Plate Purity
                r.get("S"),    # H: Colour Code
                r.get("T"),    # I: Size Code
                r.get("U"),    # J: Quanytity
            ]

            for col_idx, val in enumerate(mapped_values, start=1):
                cell = ws.cell(current_row, col_idx, val)
                apply_style(cell, template_row_styles[col_idx - 1])

            current_row += 1

    last_data_row = current_row - 1
    ws["J2"].value = f"=SUM(J4:J{last_data_row})" if last_data_row >= 4 else 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_company_orders_excel(item: dict) -> bytes:
    """
    Output 3 (per input file):
    - Uses templates/company_orders_template.xlsx (headers in A1:K1, sample items in rows 2-3,
      bordered footer block in rows 4-14)
    - Preserves all template borders, fonts, and footer labels ('TOTAL', 'Discount Value',
      'Frieght', 'Insurance', 'Total  Value') while adjusting dynamically to any row count.
    """
    wb = openpyxl.load_workbook(CO_TEMPLATE_PATH)
    ws = wb.active

    # Template has 2 sample item rows (rows 2 & 3) and footer starting at row 4
    template_item_count = 2
    item_row_styles = [snapshot_style(ws.cell(2, c)) for c in range(1, 12)]
    item_row_height = ws.row_dimensions[2].height

    # Capture the 11-row footer block (rows 4..14) values, styles, and heights intact
    footer_start_template = 4
    footer_rows_data = []
    for r_idx in range(footer_start_template, ws.max_row + 1):
        row_cells = []
        for c_idx in range(1, 12):
            cell = ws.cell(r_idx, c_idx)
            row_cells.append((cell.value, snapshot_style(cell)))
        footer_rows_data.append((ws.row_dimensions[r_idx].height, row_cells))

    # Clear everything from row 2 downward, then rebuild items + exact bordered footer
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    inv_no = item["invoice_no"]
    rows = item["rows"]
    current_row = 2

    for idx, r in enumerate(rows):
        if item_row_height:
            ws.row_dimensions[current_row].height = item_row_height

        is_last_item = idx == len(rows) - 1
        mapped_values = [
            inv_no,        # A: Q4 (Invoice)
            r.get("I"),    # B: Customer PO#
            r.get("L"),    # C: sku # (Barcode)
            r.get("K"),    # D: Design No (Stype Code)
            r.get("D"),    # E: Item (Description)
            r.get("S"),    # F: Col (Colour Code)
            r.get("F"),    # G: Gold/Plat. (Gold Plate Purity)
            r.get("U"),    # H: Total (Quanytity)
            r.get("V"),    # I: Unit Price
            r.get("W"),    # J: Total Val
            f"=SUM(J2:J{current_row})" if is_last_item else None,  # K: Total Val on last item row
        ]

        for col_idx, val in enumerate(mapped_values, start=1):
            cell = ws.cell(current_row, col_idx, val)
            apply_style(cell, item_row_styles[col_idx - 1])

        current_row += 1

    last_item_row = max(current_row - 1, 2)

    # Restore the exact bordered footer rows and update row references in formulas
    total_row = None
    discount_row = None
    freight_row = None
    insurance_row = None

    for f_height, f_cells in footer_rows_data:
        if f_height:
            ws.row_dimensions[current_row].height = f_height

        label_g = str(f_cells[6][0] or "").strip().lower()

        for col_idx, (val, st_dict) in enumerate(f_cells, start=1):
            cell = ws.cell(current_row, col_idx, val)
            apply_style(cell, st_dict)

        if label_g == "total":
            total_row = current_row
            ws.cell(current_row, 8).value = f"=SUM(H2:H{last_item_row})" if rows else 0
            ws.cell(current_row, 10).value = f"=SUM(J2:J{last_item_row})" if rows else 0
            ws.cell(current_row, 11).value = f"=SUM(J2:J{last_item_row})" if rows else 0
        elif "discount" in label_g:
            discount_row = current_row
            ws.cell(current_row, 10).value = 0
            ws.cell(current_row, 11).value = 0
        elif "frieght" in label_g or "freight" in label_g:
            freight_row = current_row
            ws.cell(current_row, 11).value = 0
        elif "insurance" in label_g:
            insurance_row = current_row
            ws.cell(current_row, 11).value = 0
        elif "total" in label_g and "value" in label_g:
            if total_row:
                ws.cell(current_row, 8).value = f"=H{total_row}"
                if discount_row and freight_row and insurance_row:
                    ws.cell(current_row, 11).value = (
                        f"=K{total_row}-K{discount_row}+K{freight_row}+K{insurance_row}"
                    )

        current_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- STREAMLIT UI ----------------
st.set_page_config(page_title="Invoice Excel Converter", layout="centered")
st.title("Invoice Excel Extractor & Formatter")

missing_templates = [
    p.name for p in (NZ_TEMPLATE_PATH, CO_TEMPLATE_PATH) if not p.exists()
]
if missing_templates:
    st.error(
        f"Missing template file(s) in `templates/` folder: {', '.join(missing_templates)}"
    )
    st.stop()

uploaded_files = st.file_uploader(
    "Upload Invoice Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.session_state.pop("zip_data", None)
    st.session_state.pop("zip_name", None)
    st.session_state.pop("summary", None)

if uploaded_files and st.button("Generate ZIP Output", type="primary"):
    parsed_files = [parse_invoice_excel(uf.read(), uf.name) for uf in uploaded_files]
    today_str = datetime.now().strftime("%d.%m.%Y")

    combined_bytes = build_combined_raw_excel(parsed_files)
    nz_bytes = build_newzealand_excel(parsed_files)

    zip_buf = io.BytesIO()
    used_names = set()

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"1_Combined_Invoices_{today_str}.xlsx", combined_bytes)
        zf.writestr(f"NewZealand Shipment Details-{today_str}.xlsx", nz_bytes)

        for idx, item in enumerate(parsed_files, start=1):
            base_label = str(item["invoice_no"] or Path(item["filename"]).stem).strip()
            co_filename = f"{base_label}-Company Orders.xlsx"

            if co_filename in used_names:
                co_filename = f"{base_label}_{idx}-Company Orders.xlsx"
            used_names.add(co_filename)

            co_bytes = build_company_orders_excel(item)
            zf.writestr(f"Company_Orders/{co_filename}", co_bytes)

    st.session_state["zip_data"] = zip_buf.getvalue()
    st.session_state["zip_name"] = f"Processed_Invoices_{today_str}.zip"
    st.session_state["summary"] = (
        f"Processed **{len(parsed_files)}** invoice(s) "
        f"({sum(len(x['rows']) for x in parsed_files)} total line items)."
    )

if "zip_data" in st.session_state:
    st.success(st.session_state["summary"])
    st.download_button(
        label="Download Output ZIP",
        data=st.session_state["zip_data"],
        file_name=st.session_state["zip_name"],
        mime="application/zip",
        type="primary",
    )