import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="NRIC Name Comparator", layout="wide")
st.title("Compare Two Vendor Lists (Full Name As Per NRIC)")

NAME_COL = "Full Name As Per NRIC"
SERIAL_COL = "S/N"

# ---------- Helpers ----------
def pick_sheet(file, key_prefix: str) -> tuple[pd.DataFrame, str]:
    ext = file.name.lower().split(".")[-1]
    engine = "openpyxl" if ext == "xlsx" else "xlrd"

    xl = pd.ExcelFile(file, engine=engine)
    sheet = st.selectbox(f"Sheet ({file.name})", xl.sheet_names, key=f"{key_prefix}_sheet")
    df = xl.parse(sheet)
    return df, sheet

def normalize_name(s: pd.Series) -> pd.Series:
    s = s.astype(str).fillna("").str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.str.upper()
    return s

def add_serial_number(df: pd.DataFrame) -> pd.DataFrame:
    df_out = df.reset_index(drop=True).copy()

    # Drop existing serial-like columns first
    serial_candidates = [
        "S/N", "SN", "SNO", "S. NO", "S. NO.", "S NO", "S NO.",
        "NO", "No", "No.", "Serial No", "Serial", "Index"
    ]
    for c in serial_candidates:
        if c in df_out.columns:
            df_out = df_out.drop(columns=[c])

    df_out.insert(0, SERIAL_COL, range(1, len(df_out) + 1))
    return df_out

def style_worksheet_exact_like_us(ws):
    """
    Match your US Cleaner styling:
    - Header fill green (94B455), bold Calibri 9
    - Borders thin, center alignment
    - Freeze top row
    - Auto-fit columns using max string length + 2 (no caps)
    - Row height 20
    """
    header_fill  = PatternFill("solid", fgColor="94B455")
    border       = Border(Side("thin"), Side("thin"), Side("thin"), Side("thin"))
    center       = Alignment(horizontal="center", vertical="center")
    normal_font  = Font(name="Calibri", size=9)
    bold_font    = Font(name="Calibri", size=9, bold=True)

    # 1) Apply borders, alignment, font to all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = center
            cell.font = normal_font

    # 2) Style header row
    for col in range(1, ws.max_column + 1):
        h = ws[f"{get_column_letter(col)}1"]
        h.fill = header_fill
        h.font = bold_font

    # 3) Freeze top row
    ws.freeze_panes = "A2"

    # 4) Auto-fit columns & set row height (EXACT same logic as your US Cleaner)
    for col in ws.columns:
        values = [len(str(cell.value)) for cell in col if cell.value is not None]
        width = max(values) if values else 10
        ws.column_dimensions[get_column_letter(col[0].column)].width = width + 2

    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 20

def build_styled_workbook_exact_like_us(sheets: dict[str, pd.DataFrame]) -> bytes:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    for sheet_name, df in sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31])

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        style_worksheet_exact_like_us(ws)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("Upload Excel A (baseline / old list)", type=["xlsx", "xls"], key="file_a")
with col2:
    file_b = st.file_uploader("Upload Excel B (new list to compare)", type=["xlsx", "xls"], key="file_b")

if file_a and file_b:
    st.subheader("1) Select sheets")
    df_a, _ = pick_sheet(file_a, "a")
    df_b, _ = pick_sheet(file_b, "b")

    # Validate required column exists
    missing = []
    if NAME_COL not in df_a.columns:
        missing.append(f"Excel A missing column: '{NAME_COL}'")
    if NAME_COL not in df_b.columns:
        missing.append(f"Excel B missing column: '{NAME_COL}'")

    if missing:
        st.error("Cannot compare because required column is missing:\n\n- " + "\n- ".join(missing))
        st.stop()

    # Compare on normalized names
    a_norm = normalize_name(df_a[NAME_COL])
    b_norm = normalize_name(df_b[NAME_COL])

    a_set = set(a_norm[a_norm != ""].tolist())
    b_set = set(b_norm[b_norm != ""].tolist())

    new_norm_set = b_set - a_set          # in B, not in A
    removed_norm_set = a_set - b_set      # in A, not in B

    # Keep FULL rows
    new_rows_b = df_b.loc[b_norm.isin(new_norm_set)].copy()
    removed_rows_a = df_a.loc[a_norm.isin(removed_norm_set)].copy()

    # Reset S/N starting from 1
    new_rows_b_out = add_serial_number(new_rows_b)
    removed_rows_a_out = add_serial_number(removed_rows_a)

    # Summary
    st.subheader("2) Summary")
    st.write(
        f"Rows A: {len(df_a)} | Rows B: {len(df_b)} | "
        f"New in Excel B: {len(new_rows_b_out)} | Removed from Excel A: {len(removed_rows_a_out)}"
    )

    # Preview (same column widths + hide index)
    st.subheader("3) Results (Preview)")
    column_cfg = {
        SERIAL_COL: st.column_config.NumberColumn(SERIAL_COL, width="small"),
        NAME_COL: st.column_config.TextColumn(NAME_COL, width="large"),
    }

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🆕 New in Excel B")
        if not new_rows_b_out.empty:
            st.dataframe(
                new_rows_b_out[[SERIAL_COL, NAME_COL]],
                use_container_width=True,
                hide_index=True,
                column_config=column_cfg
            )
        else:
            st.info("No new names found in Excel B.")

    with c2:
        st.markdown("### ❌ Removed from Excel A")
        if not removed_rows_a_out.empty:
            st.dataframe(
                removed_rows_a_out[[SERIAL_COL, NAME_COL]],
                use_container_width=True,
                hide_index=True,
                column_config=column_cfg
            )
        else:
            st.info("No names removed from Excel A.")

    # Download (EXACT same styling as US Cleaner)
    st.subheader("4) Download results")
    out_bytes = build_styled_workbook_exact_like_us({
        "New_in_Excel_B": new_rows_b_out,
        "Removed_from_Excel_A": removed_rows_a_out,
    })

    st.download_button(
        label="📥 Download comparison (Styled XLSX)",
        data=out_bytes,
        file_name="nric_name_comparison_styled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both Excel files to compare.")
