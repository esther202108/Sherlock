import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("Compare Two Vendor Lists (New Names Checker)")

# ---------- Helpers ----------
def load_excel(file, key_prefix: str):
    xl = pd.ExcelFile(file)
    sheet = st.selectbox(f"Sheet ({file.name})", xl.sheet_names, key=f"{key_prefix}_sheet")
    df = xl.parse(sheet)
    return df, sheet

def normalize_name(s: pd.Series) -> pd.Series:
    # trim, collapse multi-spaces, uppercase
    s = s.astype(str).fillna("").str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    s = s.str.upper()
    return s

def to_xlsx_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return bio.getvalue()

# ---------- UI ----------
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("Upload Excel A (baseline / old list)", type=["xlsx", "xls"], key="file_a")
with col2:
    file_b = st.file_uploader("Upload Excel B (new list to compare)", type=["xlsx", "xls"], key="file_b")

if file_a and file_b:
    st.subheader("1) Select sheets")
    df_a, sheet_a = load_excel(file_a, "a")
    df_b, sheet_b = load_excel(file_b, "b")

    st.subheader("2) Select name columns")
    name_col_a = st.selectbox("Name column in Excel A", df_a.columns, key="name_col_a")
    name_col_b = st.selectbox("Name column in Excel B", df_b.columns, key="name_col_b")

    # Optional: keep vendor column if exists (helps grouping)
    vendor_col_b = st.selectbox(
        "Optional vendor column in Excel B (for grouping)",
        ["(none)"] + list(df_b.columns),
        key="vendor_col_b"
    )
    vendor_col_b = None if vendor_col_b == "(none)" else vendor_col_b

    # ---------- Compare ----------
    a_norm = normalize_name(df_a[name_col_a])
    b_norm = normalize_name(df_b[name_col_b])

    a_set = set(a_norm[a_norm != ""].tolist())
    b_set = set(b_norm[b_norm != ""].tolist())

    # Build flagged B (new list)
    flagged_b = df_b.copy()
    flagged_b["_name_norm"] = b_norm
    flagged_b["is_new"] = ~flagged_b["_name_norm"].isin(a_set)
    flagged_b["flag"] = flagged_b["is_new"].map(lambda x: "NEW" if x else "")

    # Summaries
    new_names = sorted(list(b_set - a_set))
    removed_names = sorted(list(a_set - b_set))

    st.subheader("3) Summary")
    st.write(
        f"Rows A: {len(df_a)} | Rows B: {len(df_b)} | "
        f"Unique A: {len(a_set)} | Unique B: {len(b_set)} | "
        f"New in B: {len(new_names)} | Removed from B: {len(removed_names)}"
    )

    # Views
    view_mode = st.radio("View", ["All B (flagged)", "Only NEW in B", "New & Removed lists"], horizontal=True)

    if view_mode == "All B (flagged)":
        st.dataframe(flagged_b.drop(columns=["_name_norm"]), use_container_width=True)

    elif view_mode == "Only NEW in B":
        st.dataframe(flagged_b[flagged_b["is_new"]].drop(columns=["_name_norm"]), use_container_width=True)

        # Optional breakdown by vendor column (if provided)
        if vendor_col_b:
            st.subheader("NEW count by vendor")
            st.dataframe(
                flagged_b[flagged_b["is_new"]]
                .groupby(vendor_col_b, dropna=False)
                .size()
                .reset_index(name="new_count")
                .sort_values("new_count", ascending=False),
                use_container_width=True
            )

    else:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### New in B")
            st.dataframe(pd.DataFrame({"name": new_names}), use_container_width=True)
        with c2:
            st.markdown("### Removed from B")
            st.dataframe(pd.DataFrame({"name": removed_names}), use_container_width=True)

    # ---------- Download ----------
    st.subheader("4) Download results")
    xlsx_bytes = to_xlsx_bytes({
        "B_flagged": flagged_b.drop(columns=["_name_norm"]),
        "New_in_B": pd.DataFrame({"name": new_names}),
        "Removed_from_B": pd.DataFrame({"name": removed_names}),
    })

    st.download_button(
        "Download comparison (XLSX)",
        data=xlsx_bytes,
        file_name="vendor_name_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Upload both Excel files to compare.")
