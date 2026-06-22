"""Page 3: Review & Override Mappings — per-company view.

Each row = one company × one QB account.
Profile selector is shown at the top; saving requires a profile to be active.
"""

import io
import streamlit as st
import openpyxl
import pandas as pd
from lib.mapping_rules import map_pnl, map_bs
from lib.template_discovery import discover_template
from lib import profiles as P
from lib import db as dblib
from lib.ui import hide_streamlit_elements


hide_streamlit_elements()

st.title("🔗 Step 3 — Review & Override Mappings")

if "qb_data" not in st.session_state:
    st.warning("👈 Upload files first on **📂 Upload Files**.")
    st.stop()

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

            current_id    = st.session_state.get("active_profile_id", "")
            current_index = profile_ids.index(current_id) if current_id in profile_ids else 0

            chosen_id = st.selectbox(
                "🗂️ Active Profile",
                options=profile_ids,
                format_func=lambda x: profile_labels[profile_ids.index(x)],
                index=current_index,
                key="review_profile_selector",
                help="Select a saved profile to load + persist your mappings.",
            )
            if chosen_id != current_id:
                st.session_state.active_profile_id = chosen_id
                st.rerun()

        with pc2:
            new_profile_name = st.text_input(
                "➕ Create new profile",
                placeholder="New profile name…",
                key="review_new_profile_name",
                label_visibility="visible",
            )
        with pc3:
            st.write("")  # vertical align
            st.write("")
            if st.button("Create & activate", key="review_create_profile_btn",
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

# ── Load active profile ───────────────────────────────────────────────────────
# Auto-load most recent profile if none explicitly selected
if db_online and not st.session_state.get("active_profile_id"):
    refreshed = P.list_profiles()
    if refreshed:
        st.session_state.active_profile_id = str(refreshed[0]["_id"])

active_profile = None
saved_lookup: dict        = {}
entity_saved_lookup: dict = {}

if db_online:
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        active_profile      = P.get_profile(active_id)
        saved_lookup        = P.mapping_lookup(active_id)
        entity_saved_lookup = P.entity_mapping_lookup(active_id)

    if active_profile:
        st.success(
            f"📂 Profile: **{active_profile['name']}** — "
            f"{len(saved_lookup)} generic + {len(entity_saved_lookup)} entity-specific mappings loaded."
        )
    else:
        st.info("No active profile — session-only. Select or create a profile above to persist mappings.")

session_overrides: dict        = dict(st.session_state.get("mapping_overrides", {}))
entity_session_overrides: dict = dict(st.session_state.get("entity_mapping_overrides", {}))
generic_lookup = {**saved_lookup, **session_overrides}
entity_lookup  = {**entity_saved_lookup, **entity_session_overrides}

# ── Dynamic target lines from uploaded template ──────────────────────────────
target_lines_pnl: list = []
target_lines_bs: list  = []
template_loaded = bool(st.session_state.get("target_bytes"))

if template_loaded:
    try:
        wb_tpl = openpyxl.load_workbook(
            io.BytesIO(st.session_state.target_bytes), data_only=False)
        year_sheets = discover_template(wb_tpl)
        for stmt in [("IS", "P&L"), ("BS", "BS")]:
            stmt_key, target_list_name = stmt
            relevant = sorted([s for s in year_sheets if s.statement == stmt_key],
                              key=lambda s: -(s.year or 0))
            if not relevant:
                continue
            # Include data, preloaded, and subtotal rows — all are valid targets
            labels = [
                r.label for r in relevant[0].rows
                if r.role in ("data", "preloaded", "subtotal")
                and r.label.strip()
            ]
            if stmt_key == "IS":
                target_lines_pnl = labels
            else:
                target_lines_bs = labels
        st.caption(
            f"✅ Template dropdowns: {len(target_lines_pnl)} P&L lines, "
            f"{len(target_lines_bs)} BS lines loaded."
        )
    except Exception as e:
        st.warning(f"Could not read target template: {e}")

else:
    st.warning("⚠️ No target template uploaded — upload on **📂 Upload Files**.")


# Build target options: template labels first; fallback to auto-mapped suggestions
if target_lines_pnl or target_lines_bs:
    all_targets = sorted(set([""] + target_lines_pnl + target_lines_bs))
else:
    # No template loaded — will be filled after collect_leaves() runs
    all_targets = [""]


def collect_leaves():
    qb_data = st.session_state.qb_data
    rows = []
    for entity, info in qb_data.items():
        for stmt_kind, row_key in [("P&L", "pnl_rows"), ("BS", "bs_rows")]:
            for r in info.get(row_key, []):
                if r["is_section_only"] or r["is_total"]:
                    continue
                bc, lbl = r["breadcrumb"], r["label"]

                if stmt_kind == "P&L":
                    t, c = map_pnl(bc, lbl)
                    if t == "__SKIP__":
                        continue
                else:
                    t, c = map_bs(bc, lbl)

                entity_key  = f"E|{stmt_kind}|{entity}|{bc}|{lbl}"
                generic_key = f"{stmt_kind}|{bc}|{lbl}"

                if entity_key in entity_lookup and entity_lookup[entity_key]:
                    effective  = entity_lookup[entity_key]
                    confidence = "entity-saved"
                elif generic_key in generic_lookup and generic_lookup[generic_key]:
                    effective  = generic_lookup[generic_key]
                    confidence = "saved"
                else:
                    effective  = t or ""
                    confidence = c

                rows.append({
                    "entity_key":     entity_key,
                    "generic_key":    generic_key,
                    "Company Name":   entity,
                    "Statement":      stmt_kind,
                    "QB Account":     lbl,
                    "Breadcrumb":     bc,
                    "Auto Suggestion":t or "",
                    "Confidence":     confidence,
                    "Target Line":    effective,
                })
    return rows


rows = collect_leaves()
df   = pd.DataFrame(rows)

# ── Stats ────────────────────────────────────────────────────────────────────
total          = len(df)
entity_saved_c = (df["Confidence"] == "entity-saved").sum()
generic_saved_c= (df["Confidence"] == "saved").sum()
auto_count     = (df["Confidence"] == "auto").sum()
review_count   = (df["Confidence"] == "REVIEW").sum()
cols = st.columns(5)
cols[0].metric("Total rows",     total)
cols[1].metric("Entity-saved",   entity_saved_c)
cols[2].metric("Generic-saved",  generic_saved_c)
cols[3].metric("Auto-mapped",    auto_count,   f"{auto_count/max(total,1)*100:.0f}%")
cols[4].metric("Need review",    review_count, f"{review_count/max(total,1)*100:.0f}%")

# ── Filters ──────────────────────────────────────────────────────────────────
fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 4])
with fc1:
    filter_stmt = st.radio("Statement", ["All", "P&L only", "BS only"],
                           horizontal=True, label_visibility="collapsed")
with fc2:
    filter_conf = st.radio("Confidence", ["All", "REVIEW only", "Saved"],
                           horizontal=True, label_visibility="collapsed")
with fc3:
    qb_data_ref = st.session_state.qb_data
    companies   = ["All"] + sorted(qb_data_ref.keys())
    filter_co   = st.selectbox("Company", companies, label_visibility="collapsed")
with fc4:
    search = st.text_input("Search", "", label_visibility="collapsed",
                            placeholder="🔍 Search account...")

view = df.copy()
if filter_stmt == "P&L only":
    view = view[view["Statement"] == "P&L"]
elif filter_stmt == "BS only":
    view = view[view["Statement"] == "BS"]
if filter_conf == "REVIEW only":
    view = view[view["Confidence"] == "REVIEW"]
elif filter_conf == "Saved":
    view = view[view["Confidence"].isin(["saved", "entity-saved", "manual"])]
if filter_co != "All":
    view = view[view["Company Name"] == filter_co]
if search.strip():
    s = search.strip().lower()
    view = view[view["QB Account"].str.lower().str.contains(s, na=False) |
                view["Breadcrumb"].str.lower().str.contains(s, na=False)]

# ── Editor ───────────────────────────────────────────────────────────────────
# If no template loaded, fill target options from auto-suggestions in current view
if not all_targets or all_targets == [""]:
    auto_sugg = sorted(set(
        r["Auto Suggestion"] for _, r in df.iterrows()
        if r.get("Auto Suggestion") and r["Auto Suggestion"] not in ("", "__SKIP__")
    ))
    all_targets = sorted(set([""] + auto_sugg))

edited = st.data_editor(
    view,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    column_config={
        "entity_key":     None,
        "generic_key":    None,
        "Company Name":   st.column_config.TextColumn(disabled=True, width="medium"),
        "Statement":      st.column_config.TextColumn(disabled=True, width="small"),
        "QB Account":     st.column_config.TextColumn(disabled=True, width="medium"),
        "Breadcrumb":     st.column_config.TextColumn(disabled=True, width="large"),
        "Auto Suggestion":st.column_config.TextColumn(disabled=True, width="medium"),
        "Confidence":     st.column_config.TextColumn(disabled=True, width="small"),
        "Target Line":    st.column_config.SelectboxColumn(
            "Target Line",
            options=all_targets if len(all_targets) > 1 else [""],
            width="medium", required=False),
    },
    height=600,
    key="mapping_editor_main",
)

# ── Save ─────────────────────────────────────────────────────────────────────
st.divider()
save_col, info_col = st.columns([1, 3])
with save_col:
    save_btn = st.button("💾 Save overrides", type="primary", use_container_width=True)
with info_col:
    if active_profile:
        st.caption(f"Saves per-company to profile **{active_profile['name']}** in MongoDB.")
    elif db_online:
        st.caption("⚠️ No profile active — mappings saved to session only. Select a profile above to persist.")
    else:
        st.warning("MongoDB offline — session-only mode.")

if save_btn:
    # ── Gate: require a profile if DB is online ───────────────────────────
    if db_online and not active_profile:
        st.error(
            "❌ **No profile selected.** Please select or create a profile at the top of this page "
            "before saving — otherwise your mappings will only last for this browser session."
        )
        with st.expander("🗂️ Quick: select or create a profile now", expanded=True):
            qp_cols = st.columns([3, 2, 2])
            with qp_cols[0]:
                refreshed_profiles = P.list_profiles()
                rp_ids    = [""] + [str(p["_id"]) for p in refreshed_profiles]
                rp_labels = ["— select —"] + [p["name"] for p in refreshed_profiles]
                picked = st.selectbox("Select existing profile",
                                      options=rp_ids,
                                      format_func=lambda x: rp_labels[rp_ids.index(x)],
                                      key="save_gate_profile_select")
                if picked:
                    st.session_state.active_profile_id = picked
                    st.rerun()
            with qp_cols[1]:
                qp_name = st.text_input("Or create new", placeholder="Profile name…",
                                        key="save_gate_new_name")
            with qp_cols[2]:
                st.write("")
                st.write("")
                if st.button("Create & retry", key="save_gate_create_btn",
                             use_container_width=True):
                    if qp_name.strip():
                        nid = P.create_profile(qp_name.strip())
                        st.session_state.active_profile_id = str(nid)
                        st.success(f"Created **{qp_name.strip()}**. Click Save again.")
                        st.rerun()
                    else:
                        st.warning("Enter a profile name.")
        st.stop()

    # ── Proceed with save ─────────────────────────────────────────────────
    new_entity_overrides = dict(entity_session_overrides)
    n_saved = 0

    for _, r in edited.iterrows():
        ekey   = r["entity_key"]
        chosen = (r["Target Line"] or "").strip()
        parts  = ekey.split("|", 4)  # ["E", stmt, entity, bc, lbl]
        entity_name = parts[2] if len(parts) == 5 else ""
        stmt        = parts[1] if len(parts) >= 2 else ""
        bc          = parts[3] if len(parts) >= 4 else ""
        lbl         = parts[4] if len(parts) >= 5 else ""

        if chosen:
            new_entity_overrides[ekey] = chosen
            n_saved += 1
            if active_profile and entity_name:
                P.upsert_entity_mapping(
                    profile_id=str(active_profile["_id"]),
                    entity=entity_name, statement=stmt,
                    breadcrumb=bc, qb_account=lbl,
                    target_line=chosen, source="manual",
                )
        else:
            new_entity_overrides.pop(ekey, None)
            if active_profile and entity_name:
                try:
                    P.delete_entity_mapping(
                        str(active_profile["_id"]), entity_name, stmt, bc, lbl)
                except Exception:
                    pass

    st.session_state.entity_mapping_overrides = new_entity_overrides
    if active_profile and n_saved:
        st.success(f"✅ Saved {n_saved} per-company mappings to profile **{active_profile['name']}**.")
    else:
        st.success(f"✅ Saved {len(new_entity_overrides)} entity overrides in session.")

if template_loaded:
    st.info("👉 Next: **💾 Generate Linked Workbook**.")
else:
    st.warning("⚠️ No target template uploaded. Go to **📂 Upload Files**.")
