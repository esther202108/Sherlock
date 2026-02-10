st.subheader("3) Results (Preview)")
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
