import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ─────────────────────────────────────────────────────────────
# Streamlit setup
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="NRIC Name Comparator", layout="wide")
st.title("Compare Two Vendor Lists (Full Name As Per NRIC)")

NAME_COL = "Full Name As Per NRIC"
SERIAL_COL = "S/N"

# Fixed column widths (EXACT match to US Cleaner)
# NOTE: This works reliably only when your output columns align with A..M positions.
COLUMN_WIDTHS = {
    "A": 3.38,   # S/N
    "C": 23.06,
    "D": 25,
    "E": 17.63,
    "F": 26.25,
    "G": 13.94,
    "H": 24.06,
    "I": 18.38,
    "J": 20.31,
    "K": 4,
    "L": 5.81,
    "M": 11.5,
}

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
def pick_sheet(file, key_prefix: str):
    ext = file.name.lower().split(".")[-1]
    engine = "openpyxl" if ext == "xlsx" else "xlrd"
    xl = pd.ExcelFile(file, engine=engine)
    sheet = st.selectbox(f"Sheet ({file.name})", xl.sheet_names, key=f"{key_prefix}_sheet")
    return xl.parse(sheet)

def normalize_name(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .fillna("")
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.upper()
    )

def add_serial_number(df: pd.DataFrame) -> pd.DataFrame:
    df = df.reset_index(drop=True).copy()

    # Remove any existing serial column (common variations)
    for c in ["S/N", "SN", "SNO", "S. NO", "S. NO.", "S NO", "S NO.", "No", "No.", "Index", "Serial", "Serial No"]:
        if c in df.columns:
            df.drop(columns=[c], inplace=True)

    df.insert(0, SERIAL_COL, range(1, len(df) + 1))
    return df

# ─────────────────────────────────────────────────────────────
# Excel styling (MATCH US CLEANER)
# ─────────────────────────────────────────────────────────────
def apply_us_style(ws):
    header_fill = PatternFill("solid", fgColor="94B455")
    border = Border(Side("thin"), Side("thin"), Side("thin"), Side("thin"))
    center = Alignment(horizontal="center", vertical="center")
    normal_font = Font(name="Calibri", size=9)
    bold_font = Font(name="Calibri", size=9, bold=True)

    # Borders, alignment, font
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = center
            cell.font = normal_font

    # Header styling
    for col in range(1, ws.max_column + 1):
        cell = ws[f"{get_column_letter(col)}1"]
        cell.fill = header_fill
        cell.font = bold_font

    # Freeze header
    ws.freeze_panes = "A2"

    # Row height
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 20

def apply_fixed_column_widths(ws):
    """
    IMPORTANT FIX:
    - Set widths UNCONDITIONALLY (do not check membership).
    - Also set S/N (A) explicitly even if not in the mapping.
    """
    # Always force S/N column narrower
    ws.column_dimensions["A"].width = COLUMN_WIDTHS.get("A", 3.38)

    # Apply the rest
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

def build_workbook(sheets: dict[str, pd.DataFrame]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    for name, df in sheets.items():
        ws = wb.create_sheet(title=name[:31])

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        apply_us_style(ws)
        apply_fixed_column_widths(ws)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────
c1, c2 = st.columns(2)
with c1:
    file_a = st.file_uploader("Upload Excel A (baseline)", type=["xlsx", "xls"])
with c2:
    file_b = st.file_uploader("Upload Excel B (new list)", type=["xlsx", "xls"])

if file_a and file_b:
    st.subheader("1) Select sheets")
    df_a = pick_sheet(file_a, "a")
    df_b = pick_sheet(file_b, "b")

    if NAME_COL not in df_a.columns or NAME_COL not in df_b.columns:
        st.error(f"Both files must contain column: '{NAME_COL}'")
        st.stop()

    a_norm = normalize_name(df_a[NAME_COL])
    b_norm = normalize_name(df_b[NAME_COL])

    # Remove blanks before set operations (prevents '' being treated as a name)
    a_set = set(a_norm[a_norm != ""].tolist())
    b_set = set(b_norm[b_norm != ""].tolist())

    new_set = b_set - a_set
    removed_set = a_set - b_set

    new_rows = df_b.loc[b_norm.isin(new_set)].copy()
    removed_rows = df_a.loc[a_norm.isin(removed_set)].copy()

    new_out = add_serial_number(new_rows)
    removed_out = add_serial_number(removed_rows)

    # Summary
    st.subheader("2) Summary")
    st.write(
        f"New in Excel B: {len(new_out)} | "
        f"Removed from Excel A: {len(removed_out)}"
    )

    # Preview (tight S/N)
    st.subheader("3) Results (Preview)")
    col_cfg = {
        SERIAL_COL: st.column_config.NumberColumn(SERIAL_COL, width="xsmall"),
        NAME_COL: st.column_config.TextColumn(NAME_COL, width="large"),
    }

    l, r = st.columns(2)
    with l:
        st.markdown("### 🆕 New in Excel B")
        if not new_out.empty:
            st.dataframe(
                new_out[[SERIAL_COL, NAME_COL]],
                hide_index=True,
                use_container_width=True,
                column_config=col_cfg,
            )
        else:
            st.info("No new names found in Excel B.")

    with r:
        st.markdown("### ❌ Removed from Excel A")
        if not removed_out.empty:
            st.dataframe(
                removed_out[[SERIAL_COL, NAME_COL]],
                hide_index=True,
                use_container_width=True,
                column_config=col_cfg,
            )
        else:
            st.info("No names removed from Excel A.")

    # Download
    st.subheader("4) Download results")
    output = build_workbook({
        "New_in_Excel_B": new_out,
        "Removed_from_Excel_A": removed_out,
    })

    st.download_button(
        "📥 Download comparison (Styled XLSX)",
        data=output,
        file_name="nric_name_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Upload both Excel files to begin.")
