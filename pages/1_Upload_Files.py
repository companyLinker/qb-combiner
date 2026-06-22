"""Page 1: Upload QB exports + (optionally) the target combination template."""

import io
import re
import streamlit as st
import pandas as pd
import openpyxl

from lib.parser import parse_uploaded_files
from lib.template_discovery import discover_template
from lib.ui import hide_streamlit_elements
from lib import db as dblib
from lib import profiles as P

hide_streamlit_elements()

st.title("📂 Step 1 — Upload Files")

# ── Profile selector at the TOP ───────────────────────────────────────────────
db_online = dblib.is_connected()

if db_online:
    all_profiles = P.list_profiles()
    with st.container(border=True):
        pc1, pc2, pc3 = st.columns([3, 2, 2])
        with pc1:
            profile_options = {str(p["_id"]): p["name"] for p in all_profiles}
            profile_ids     = [""] + list(profile_options.keys())
            profile_labels  = ["— select a profile —"] + list(profile_options.values())
            current_id      = st.session_state.get("active_profile_id", "")
            current_index   = profile_ids.index(current_id) if current_id in profile_ids else 0
            chosen_id = st.selectbox(
                "🗂️ Active Profile",
                options=profile_ids,
                format_func=lambda x: profile_labels[profile_ids.index(x)],
                index=current_index,
                key="p1_profile_selector",
                help="Select a profile so your mapping decisions are saved and reloaded across sessions.",
            )
            if chosen_id != current_id:
                st.session_state.active_profile_id = chosen_id
                st.rerun()
        with pc2:
            new_profile_name = st.text_input(
                "➕ Create new profile",
                placeholder="New profile name…",
                key="p1_new_profile_name",
            )
        with pc3:
            st.write("")
            st.write("")
            if st.button("Create & activate", key="p1_create_profile_btn",
                         use_container_width=True, type="secondary"):
                if new_profile_name.strip():
                    new_id = P.create_profile(new_profile_name.strip())
                    st.session_state.active_profile_id = str(new_id)
                    st.success(f"✅ Created profile **{new_profile_name.strip()}**.")
                    st.rerun()
                else:
                    st.warning("Enter a name first.")
else:
    st.warning(f"MongoDB unavailable — session-only mode. ({dblib.connection_error()})")

# Auto-load most recent profile if none selected
if db_online and not st.session_state.get("active_profile_id"):
    refreshed_profiles = P.list_profiles()
    if refreshed_profiles:
        st.session_state.active_profile_id = str(refreshed_profiles[0]["_id"])

# Resolve active profile for display
active_profile = None
if db_online:
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        active_profile = P.get_profile(active_id)
    if active_profile:
        st.success(f"📂 Active profile: **{active_profile['name']}**")
    else:
        st.info("No profile active — session-only. Select or create a profile above to persist mappings.")

st.divider()

# ── QB File Upload ─────────────────────────────────────────────────────────────
st.markdown("Upload your **QuickBooks Excel exports** (one .xlsx per entity / LLC).")
st.markdown("Each file should have a *Profit & Loss* sheet AND a *Balance Sheet As of …* sheet.")

uploaded = st.file_uploader(
    "Drop QuickBooks .xlsx files here (multi-select)",
    type=["xlsx"],
    accept_multiple_files=True,
    key="qb_uploader",
)

if uploaded:
    with st.spinner(f"Parsing {len(uploaded)} files…"):
        qb_data = parse_uploaded_files(uploaded)
    st.session_state.qb_data = qb_data
    st.success(f"✅ Parsed {len(qb_data)} entities.")

    rows = []
    for ent, info in qb_data.items():
        variant = info.get("variant", "single")
        variant_label = "Dual-year (CY+PY)" if variant == "triple" else "Single-year"
        rows.append({
            "Entity": ent,
            "File": info.get("file", ""),
            "Format": variant_label,
            "P&L Period (CY)": info.get("pnl_period_cy") or info.get("pnl_period", ""),
            "P&L Period (PY)": info.get("pnl_period_py", "") or "",
            "BS Period (CY)": info.get("bs_period_cy") or info.get("bs_period", ""),
            "P&L Rows": len(info.get("pnl_rows", [])),
            "BS Rows": len(info.get("bs_rows", [])),
            "Error": info.get("error", ""),
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)


st.divider()
st.markdown("### Target Combination Template *(optional but recommended)*")
st.markdown(
    "Upload your **AP NE COMBINATION** (or equivalent) target template — the workbook "
    "with entity columns across and target chart-of-accounts down column A. "
    "Without this you can still produce the consolidated master and CoA variants, "
    "but not the final SUMIFS-linked workbook."
)

target = st.file_uploader(
    "Drop target template .xlsx here",
    type=["xlsx"],
    accept_multiple_files=False,
    key="target_uploader",
)

if target:
    st.session_state.target_bytes = target.getvalue()
    st.session_state.target_filename = target.name
    # Clear old configuration when a new template is uploaded
    st.session_state.pop("selected_template_sheets", None)
    st.session_state.pop("template_entity_mapping", None)
    st.success(f"✅ Target template loaded: **{target.name}**")


# ── Template Configuration (tab selection + entity mapping) ─────────────────
def _norm(s: str) -> str:
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _auto_match(template_col: str, qb_entities: list) -> str | None:
    """Return best-matching QB entity name for a template column, or None."""
    norm_t = _norm(template_col)
    # Exact normalized match first
    for qb_e in qb_entities:
        if _norm(qb_e) == norm_t:
            return qb_e
    # Word-overlap fuzzy: ≥60% word overlap
    t_words = set(norm_t.split())
    best, best_score = None, 0.0
    for qb_e in qb_entities:
        q_words = set(_norm(qb_e).split())
        union = t_words | q_words
        if not union:
            continue
        score = len(t_words & q_words) / len(union)
        if score > best_score and score >= 0.60:
            best_score, best = score, qb_e
    return best


if st.session_state.get("target_bytes"):
    st.divider()
    st.markdown("### ⚙️ Template Configuration")

    try:
        wb_tpl = openpyxl.load_workbook(
            io.BytesIO(st.session_state.target_bytes), data_only=False
        )
        all_sheet_names = wb_tpl.sheetnames
        year_sheets_all = discover_template(wb_tpl)
        classified = {ys.sheet_name: ys for ys in year_sheets_all}

        # ── a) Sheet Selection ─────────────────────────────────────────────
        st.markdown("**📋 Select which template sheets to include in the output:**")

        default_selected = st.session_state.get(
            "selected_template_sheets",
            [sn for sn in all_sheet_names if sn in classified],
        )
        # Validate defaults against current template
        default_selected = [s for s in default_selected if s in all_sheet_names]

        selected_sheets = st.multiselect(
            "Template sheets",
            options=all_sheet_names,
            default=default_selected,
            label_visibility="collapsed",
            help="IS/BS classified sheets get SUMIFS formulas written. "
                 "Other sheets are preserved as-is in the output.",
            key="sheet_selector",
        )
        st.session_state.selected_template_sheets = selected_sheets

        if selected_sheets:
            info_rows = []
            for sn in selected_sheets:
                if sn in classified:
                    ys = classified[sn]
                    info_rows.append({
                        "Sheet": sn,
                        "Type": ys.statement,
                        "Year": str(ys.year) if ys.year else "—",
                        "Entity Cols": len(ys.entity_cols),
                        "Data Rows": sum(1 for r in ys.rows if r.role == "data"),
                    })
                else:
                    info_rows.append({
                        "Sheet": sn, "Type": "OTHER",
                        "Year": "—", "Entity Cols": "—", "Data Rows": "—",
                    })
            st.dataframe(pd.DataFrame(info_rows), use_container_width=True, hide_index=True)

        # ── b) Entity → QB Company Mapping ────────────────────────────────
        selected_ys = [classified[sn] for sn in selected_sheets if sn in classified]

        if selected_ys:
            # Collect unique entity column names (order-preserving)
            seen_cols: dict[str, bool] = {}
            for ys in selected_ys:
                for _col, name in ys.entity_cols:
                    seen_cols[name] = True
            unique_template_cols = list(seen_cols.keys())

            if unique_template_cols:
                st.markdown("---")
                st.markdown("**🔗 Map QB Companies → Template Columns**")

                qb_data_now = st.session_state.get("qb_data", {})
                qb_entities = list(qb_data_now.keys())

                if not qb_entities:
                    st.info("ℹ️ Upload QB files above first to configure entity mapping.")
                else:
                    st.caption(
                        "For each template column, choose which uploaded QB company's data "
                        "should flow into it. Set to **'— Skip —'** to leave that column untouched."
                    )

                    SKIP = "— Skip (leave blank) —"
                    options = [SKIP] + qb_entities
                    existing_mapping: dict = st.session_state.get("template_entity_mapping", {})

                    # Header row
                    hcol1, hcol2 = st.columns([5, 5])
                    hcol1.markdown("**Template Column**")
                    hcol2.markdown("**QB Company Data**")
                    st.markdown(
                        "<hr style='margin:4px 0 8px 0; border-color:#e0e0e0;'>",
                        unsafe_allow_html=True,
                    )

                    new_mapping: dict[str, str | None] = {}
                    for i, tc in enumerate(unique_template_cols):
                        # Determine default
                        if tc in existing_mapping:
                            default_entity = existing_mapping[tc]
                        else:
                            default_entity = _auto_match(tc, qb_entities)

                        default_option = (
                            default_entity
                            if default_entity and default_entity in options
                            else SKIP
                        )

                        col_label, col_select = st.columns([5, 5])
                        with col_label:
                            # Color-code: green = auto/saved match, grey = skip
                            color = "#1a6b2e" if default_option != SKIP else "#888"
                            st.markdown(
                                f"<div style='padding:6px 0; font-size:0.88em; color:{color};'>"
                                f"<b>{tc}</b></div>",
                                unsafe_allow_html=True,
                            )
                        with col_select:
                            chosen = st.selectbox(
                                label=f"col_{i}",
                                options=options,
                                index=options.index(default_option),
                                label_visibility="collapsed",
                                key=f"temap_{i}",
                            )
                        new_mapping[tc] = None if chosen == SKIP else chosen

                    st.markdown("<br>", unsafe_allow_html=True)
                    save_col, info_col = st.columns([2, 5])
                    with save_col:
                        save_cfg_btn = st.button(
                            "💾 Save Configuration", type="primary", use_container_width=True,
                            key="p1_save_config_btn"
                        )
                    with info_col:
                        n_auto = sum(
                            1 for tc in unique_template_cols
                            if tc not in existing_mapping and _auto_match(tc, qb_entities)
                        )
                        if active_profile:
                            st.caption(
                                f"{len(unique_template_cols)} unique template columns • "
                                f"{n_auto} auto-matched • adjust then click Save. "
                                f"Saving to profile **{active_profile['name']}**."
                            )
                        elif db_online:
                            st.caption(
                                f"{len(unique_template_cols)} unique template columns • "
                                f"{n_auto} auto-matched • ⚠️ No profile active — "
                                "select one above to persist."
                            )
                        else:
                            st.caption(
                                f"{len(unique_template_cols)} unique template columns • "
                                f"{n_auto} auto-matched • adjust then click Save."
                            )

                    if save_cfg_btn:
                        # ── Gate: require a profile if DB is online ────────
                        if db_online and not active_profile:
                            st.error(
                                "❌ **No profile selected.** Please select or create a profile "
                                "at the top of this page before saving — otherwise your "
                                "configuration will only last for this browser session."
                            )
                            with st.expander("🗂️ Quick: select or create a profile now",
                                             expanded=True):
                                qp_cols = st.columns([3, 2, 2])
                                with qp_cols[0]:
                                    refreshed_profiles = P.list_profiles()
                                    rp_ids    = [""] + [str(p["_id"]) for p in refreshed_profiles]
                                    rp_labels = ["— select —"] + [p["name"] for p in refreshed_profiles]
                                    picked = st.selectbox(
                                        "Select existing profile",
                                        options=rp_ids,
                                        format_func=lambda x: rp_labels[rp_ids.index(x)],
                                        key="p1_save_gate_select"
                                    )
                                    if picked:
                                        st.session_state.active_profile_id = picked
                                        st.rerun()
                                with qp_cols[1]:
                                    qp_name = st.text_input(
                                        "Or create new", placeholder="Profile name…",
                                        key="p1_save_gate_new_name"
                                    )
                                with qp_cols[2]:
                                    st.write("")
                                    st.write("")
                                    if st.button("Create & retry", key="p1_save_gate_create_btn",
                                                 use_container_width=True):
                                        if qp_name.strip():
                                            nid = P.create_profile(qp_name.strip())
                                            st.session_state.active_profile_id = str(nid)
                                            st.success(
                                                f"Created **{qp_name.strip()}**. Click Save again."
                                            )
                                            st.rerun()
                                        else:
                                            st.warning("Enter a profile name.")
                        else:
                            # ── Save configuration ─────────────────────────
                            st.session_state.template_entity_mapping = new_mapping
                            n_mapped = sum(1 for v in new_mapping.values() if v)
                            st.success(
                                f"✅ Saved: **{len(selected_sheets)} sheets** selected, "
                                f"**{n_mapped}/{len(unique_template_cols)}** columns mapped."
                            )

    except Exception as e:
        st.error(f"Could not process template for configuration: {e}")


# Next-step hint
if st.session_state.get("qb_data"):
    st.divider()
    st.info("👉 Next: go to **📊 Variants & Analysis** in the sidebar.")
