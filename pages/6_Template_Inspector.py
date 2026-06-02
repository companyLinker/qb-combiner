"""Page 6: Template Inspector — preview the discovered structure of the target template.

Lets the user verify which sheets, entities, rows, and existing formulas the
builder sees BEFORE running the generation. Catches typos, missed entities,
unexpected subtotal layouts, etc. before they bake into output.
"""

import io
import streamlit as st
import openpyxl
from lib.template_discovery import discover_template, summarize
from lib.ui import hide_streamlit_elements


if st.secrets.get("APP_PASSWORD", None) and not st.session_state.get("authed"):
    st.warning("Please sign in from the home page first.")
    st.stop()

hide_streamlit_elements()


st.title("Template Inspector")

st.markdown(
    "Preview what the builder sees in your target combination template. "
    "Use this to verify entity columns, subtotal rows, and cross-sheet formulas "
    "are correctly identified before generating."
)

if "target_bytes" not in st.session_state:
    st.warning("Upload a target template first on Upload Files.")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(st.session_state.target_bytes), data_only=False)
year_sheets = discover_template(wb)

if not year_sheets:
    st.error(f"No BS/IS sheets recognized. Sheets present: {wb.sheetnames}")
    st.stop()

st.success(f"{len(year_sheets)} year-sheet(s) discovered.")

# Top-level summary
totals = {
    "Year-sheets": len(year_sheets),
    "Total data rows": sum(sum(1 for r in s.rows if r.role == "data") for s in year_sheets),
    "Total subtotals preserved": sum(sum(1 for r in s.rows if r.role == "subtotal") for s in year_sheets),
    "Total cross-sheet refs preserved": sum(sum(1 for r in s.rows if r.role == "cross_ref") for s in year_sheets),
    "Total preloaded cells": sum(sum(1 for r in s.rows if r.role == "preloaded") for s in year_sheets),
}
cols = st.columns(len(totals))
for i, (k, v) in enumerate(totals.items()):
    cols[i].metric(k, v)


# Per-sheet detail
st.divider()
for ys in year_sheets:
    with st.expander(f"{ys.sheet_name} — {ys.statement} {ys.year or '(no year)'} — "
                     f"{len(ys.entity_cols)} entities, {len(ys.rows)} rows"):
        c1, c2, c3 = st.columns(3)
        c1.metric("Data rows (writable)", sum(1 for r in ys.rows if r.role == "data"))
        c2.metric("Subtotals (preserved)", sum(1 for r in ys.rows if r.role == "subtotal"))
        c3.metric("Cross-refs (preserved)", sum(1 for r in ys.rows if r.role == "cross_ref"))

        st.markdown("**Entity columns**")
        st.dataframe(
            [{"Column": chr(64 + c) if c < 27 else "AB"[c - 27], "Header": h}
             for c, h in ys.entity_cols],
            use_container_width=True, hide_index=True,
        )

        if ys.summary_cols:
            st.markdown("**Summary columns**")
            st.dataframe(
                [{"Column": chr(64 + c) if c < 27 else "A" + chr(64 + c - 26), "Role": r}
                 for c, r in ys.summary_cols],
                use_container_width=True, hide_index=True,
            )

        st.markdown("**Row breakdown**")
        st.dataframe(
            [
                {
                    "Row": r.row_idx,
                    "Label": r.label,
                    "Role": r.role,
                    "Existing formula": (r.existing_formula or "")[:60],
                }
                for r in ys.rows
            ],
            use_container_width=True, hide_index=True,
            height=min(400, 50 + len(ys.rows) * 26),
        )
