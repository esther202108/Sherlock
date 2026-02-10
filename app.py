import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="NRIC Name Comparator", layout="wide")
st.title("Compare Two Vendor Lists (Full Name As Per NRIC)")

NAME_COL = "Full Name As Per NRIC"

# ---------- Helpers ----------
def pick_sheet(file, key_prefix: str) -> tuple[pd.DataFrame, str]:
    ext = file.name.lower().split(".")[-1]
    engine = "openpyxl" if ext == "xlsx" else "xlrd"

    xl = pd.ExcelFile(file, engine=engine)
    sheet = st.selectbox(
        f"Sheet ({file.name})",
        xl.sheet_names,
        key=f"{key_prefix}_sheet"
    )
    df = xl.parse(sheet)
    return df, sheet

def normalize_name(s: pd.Series) -> pd.Series:
    s = s.astype(str).fillna("").str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.str.upper()
    return s

def to_xlsx_bytes(df_dict: dict) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=str(sheet_name)[:31])
    return bio.getvalue()

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader(
        "Upload Excel A (baseline / old list)",
        type=["xlsx", "xls"],
        key="file_a"
    )
with col2:
    file_b = st.file_uploader(
        "Upload Excel B (new list to compare)",
        type=["xlsx", "xls"],
        key="file_b"
    )

if file_a and file_b:
    st.subheader("1) Select sheets")
    df_a, sheet_a = pick_sheet(file_a, "a")
    df_b, sheet_b = pick_sheet(file_b, "b")

    # ---------- Validate column ----------
    missing = []
    if NAME_COL not in df_a.columns:
        missing.append(f"Excel A missing column: '{NAME_COL}'")
    if NAME_COL not in df_b.columns:
        missing.append(f"Excel B missing column: '{NAME_COL}'")

    if missing:
        st.error("Cannot compare because required column is missing:\n\n- " + "\n- ".join(missing))
        st.stop()

    # ---------- Compare ----------
    a_norm = normalize_name(df_a[NAME_COL])
    b_norm = normalize_name(df_b[NAME_COL])

    a_set = set(a_norm[a_norm != ""].tolist())
    b_set = set(b_norm[b_norm != ""].tolist())

    new_names = sorted(list(b_set - a_set))
    removed_names = sorted(list(a_set - b_set))

    # ---------- Summary ----------
    st.subheader("2) Summary")
    st.write(
        f"Rows A: {len(df_a)} | Rows B: {len(df_b)} | "
        f"Unique A: {len(a_set)} | Unique B: {len(b_set)} | "
        f"New in B: {len(new_names)} | Removed from B: {len(removed_names)}"
    )

    # ---------- Views ----------
    st.subheader("3) Results")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 🆕 New in Excel B")
        if new_names:
            st.dataframe(
                pd.DataFrame({NAME_COL: new_names}),
                use_container_width=True
            )
        else:
            st.info("No new names found.")

    with c2:
        st.markdown("### ❌ Removed from Excel B")
        if removed_names:
            st.dataframe(
                pd.DataFrame({NAME_COL: removed_names}),
                use_container_width=True
            )
        else:
            st.info("No removed names found.")

    # ---------- Download ----------
    st.subheader("4) Download results")
    xlsx_bytes = to_xlsx_bytes({
        "New_in_B": pd.DataFrame({NAME_COL: new_names}),
        "Removed_from_B": pd.DataFrame({NAME_COL: removed_names}),
    })

    st.download_button(
        "Download comparison (XLSX)",
        data=xlsx_bytes,
        file_name="nric_name_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both Excel files to compare.")
