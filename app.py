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

    # drop existing serial-like columns to avoid "cannot insert already exists"
    serial_candidates = ["S/N", "SN", "SNO", "S. NO", "S. NO.", "S NO", "S NO.", "NO", "No", "No.", "Serial No", "Serial", "Index"]
    for c in serial_candidates:
        if c in df_out.columns:
            df_out = df_out.drop(columns=[c])

    df_out.insert(0, SERIAL_COL, range(1, len(df_out) + 1))
    return df_out

def style_worksheet_like_us(ws):
    # Same “look” as your US cleaner
    header_fill  = PatternFill("solid", fgColor="94B455")
    border       = Border(Side("thin"), Side("thin"), Side("thin"), Side("thin"))
    center       = Alignment(horizontal="center", vertical="center", wrap_text=True)
    normal_font  = Font(name="Calibri", size=9)
    bold_font    = Font(name="Calibri", size=9, bold=True)

    # Apply borders/alignment/font to all cells
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = center
            cell.font = normal_font

    # Style header row
    for col in range(1, ws.max_column + 1):
        h = ws[f"{get_column_letter(col)}1"]
        h.fill = header_fill
        h.font = bold_font

    # Freeze top row
    ws.freeze_panes = "A2"

    # Auto-fit columns & set row height
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        width = max(max_len, 10) + 2
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(width, 60)  # cap to avoid crazy-wide cols

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 20

def build_styled_workbook(sheets: dict[str, pd.DataFrame]) -> bytes:
    wb = Workbook()
    # Remove default sheet
    default = wb.active
    wb.remove(default)

    for sheet_name, df in sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31])

        # Write dataframe (header + rows)
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        # Apply style
        style_worksheet_like_us(ws)

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

    # Valida
