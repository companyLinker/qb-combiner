"""Page 4: Generate the SUMIFS-linked combination workbook and download it."""

import streamlit as st
import pandas as pd
from lib.linked_builder import build_linked_workbook, HAVE_RAPIDFUZZ
from lib import profiles as P
from lib import db as dblib


if st.secrets.get("APP_PASSWORD", None) and not st.session_state.get("authed"):
    st.warning("Please sign in from the home page first.")
    st.stop()


st.title("💾 Step 4 — Generate Linked Workbook")

if "qb_data" not in st.session_state:
    st.warning("Upload files first on **📂 Upload Files**.")
    st.stop()
if "target_bytes" not in st.session_state:
    st.warning("No target template uploaded. Upload it on **📂 Upload Files**.")
    st.stop()

qb_data       = st.session_state.qb_data
target_bytes  = st.session_state.target_bytes

# ── Profile / overrides ──────────────────────────────────────────────────────
profile_lookup = {}
entity_lookup  = {}
active_profile = None
if dblib.is_connected():
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        active_profile = P.get_profile(active_id)
        profile_lookup = P.mapping_lookup(active_id)
        entity_lookup  = P.entity_mapping_lookup(active_id)

session_overrides        = st.session_state.get("mapping_overrides", {})
entity_session_overrides = st.session_state.get("entity_mapping_overrides", {})
combined        = {**profile_lookup, **session_overrides}
entity_combined = {**entity_lookup, **entity_session_overrides}

# ── Template configuration from Upload Files page ────────────────────────────
selected_sheets    = st.session_state.get("selected_template_sheets")
entity_col_mapping = st.session_state.get("template_entity_mapping")

# ── Summary metrics ──────────────────────────────────────────────────────────
cols = st.columns(5)
cols[0].metric("Entities",          len(qb_data))
cols[1].metric("Profile",           active_profile["name"] if active_profile else "—")
cols[2].metric("From profile",      len(profile_lookup))
cols[3].metric("Session overrides", len(session_overrides))
cols[4].metric("Entity overrides",  len(entity_session_overrides))

st.caption(
    f"Fuzzy matching: {'enabled (rapidfuzz)' if HAVE_RAPIDFUZZ else 'fallback (exact-normalized only)'}"
)

# ── Template configuration status ────────────────────────────────────────────
st.divider()
st.markdown("### 📋 Template Configuration")

cfg_col1, cfg_col2 = st.columns(2)

with cfg_col1:
    st.markdown("**Selected Sheets**")
    if selected_sheets:
        for sn in selected_sheets:
            st.markdown(f"- `{sn}`")
    else:
        st.info("No sheet selection saved — all IS/BS sheets will be processed.")
        st.caption("Go to **📂 Upload Files** → Template Configuration to select sheets.")

with cfg_col2:
    st.markdown("**Entity Column Mapping**")
    if entity_col_mapping:
        map_rows = []
        for tc, qb_e in entity_col_mapping.items():
            map_rows.append({
                "Template Column": tc,
                "QB Company": qb_e if qb_e else "— Skip —",
                "Status": "✅ Mapped" if qb_e else "⏭️ Skip",
            })
        st.dataframe(
            pd.DataFrame(map_rows),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Template Column": st.column_config.TextColumn(width="large"),
                "QB Company":      st.column_config.TextColumn(width="large"),
                "Status":          st.column_config.TextColumn(width="small"),
            },
        )
        n_mapped = sum(1 for v in entity_col_mapping.values() if v)
        st.caption(f"{n_mapped}/{len(entity_col_mapping)} columns mapped.")
    else:
        st.info(
            "No entity mapping saved — template column headers will be matched "
            "verbatim to QB entity names."
        )
        st.caption("Go to **📂 Upload Files** → Template Configuration to set up mapping.")

st.divider()

# ── Options ──────────────────────────────────────────────────────────────────
overwrite = st.checkbox(
    "Overwrite preloaded values (historical year-sheets)",
    value=False,
    help="Multi-year templates often contain prior-year actuals as static numbers. "
         "Unchecked = preserve them; checked = recompute from QB.",
)

# ── Generate ─────────────────────────────────────────────────────────────────
if st.button("🚀 Generate Linked Workbook", type="primary"):
    with st.spinner("Discovering template structure and writing SUMIFS…"):
        buf, mapping, year_sheets, report = build_linked_workbook(
            qb_data,
            target_bytes,
            mapping_overrides=combined,
            entity_mapping_overrides=entity_combined,
            overwrite_preloaded=overwrite,
            selected_sheets=selected_sheets or None,
            entity_col_mapping=entity_col_mapping or None,
        )
    st.session_state.linked_buf          = buf.getvalue()
    st.session_state.linked_mapping      = mapping
    st.session_state.linked_year_sheets  = year_sheets
    st.session_state.linked_report       = report
    st.success("✅ Workbook generated!")

    if active_profile:
        n_auto   = sum(1 for v in mapping.values() if v[1] == "auto")
        n_manual = sum(1 for v in mapping.values() if v[1] == "manual")
        n_review = sum(1 for v in mapping.values() if v[1] == "REVIEW")
        try:
            P.log_run(
                profile_id=str(active_profile["_id"]),
                entities_count=len(qb_data),
                auto_count=n_auto, manual_count=n_manual, review_count=n_review,
                output_filename="02_LINKED_Combination.xlsx",
            )
        except Exception as e:
            st.warning(f"Workbook built but run logging failed: {e}")

# ── Download ─────────────────────────────────────────────────────────────────
if "linked_buf" in st.session_state:
    st.download_button(
        "⬇ Download Linked Combination Workbook",
        st.session_state.linked_buf,
        "02_LINKED_Combination.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    report = st.session_state.linked_report
    rc1, rc2, rc3, rc4 = st.columns(4)
    rc1.metric("Sheets processed",     report["n_year_sheets"])
    rc2.metric("Cells written",        report["cells_written"])
    rc3.metric("Subtotals preserved",  report["cells_skipped_subtotal"])
    rc4.metric("Cross-refs preserved", report["cells_skipped_crossref"])

    if report["cells_skipped_preloaded"]:
        st.info(
            f"{report['cells_skipped_preloaded']} preloaded cells left intact "
            "(check overwrite box above to recompute)."
        )

    if report["rows_unmapped"]:
        with st.expander(f"⚠️ Rows with no QB target match ({len(report['rows_unmapped'])})"):
            st.dataframe(report["rows_unmapped"], use_container_width=True, hide_index=True)
