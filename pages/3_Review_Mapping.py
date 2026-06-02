"""Page 3: Review & Override Mappings — per-company view.

Each row = one company × one QB account.
Company Name column added before Statement.
"""

import io
import streamlit as st
import openpyxl
import pandas as pd
from lib.mapping_rules import map_pnl, map_bs
from lib.template_discovery import discover_template
from lib import profiles as P
from lib import db as dblib


if st.secrets.get("APP_PASSWORD", None) and not st.session_state.get("authed"):
    st.warning("Please sign in from the home page first.")
    st.stop()

st.title("🔗 Step 3 — Review & Override Mappings")

if "qb_data" not in st.session_state:
    st.warning("👈 Upload files first on **📂 Upload Files**.")
    st.stop()

# ── Auto-load most recent profile if none selected ──────────────────────────
if dblib.is_connected() and not st.session_state.get("active_profile_id"):
    all_profiles = P.list_profiles()
    if all_profiles:
        st.session_state.active_profile_id = str(all_profiles[0]["_id"])

# ── Active profile & lookups ─────────────────────────────────────────────────
active_profile = None
saved_lookup: dict        = {}
entity_saved_lookup: dict = {}
if dblib.is_connected():
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        active_profile      = P.get_profile(active_id)
        saved_lookup        = P.mapping_lookup(active_id)
        entity_saved_lookup = P.entity_mapping_lookup(active_id)

    if active_profile:
        st.success(f"📂 Profile: **{active_profile['name']}** — "
                   f"{len(saved_lookup)} generic + {len(entity_saved_lookup)} entity-specific mappings loaded.")
    else:
        st.info("No active profile. Go to **🗂️ Profiles** to create or pick one.")
else:
    st.warning(f"MongoDB unavailable — session-only mode. ({dblib.connection_error()})")

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
        for stmt in ["IS", "BS"]:
            relevant = sorted([s for s in year_sheets if s.statement == stmt],
                              key=lambda s: -(s.year or 0))
            if not relevant:
                continue
            labels = [r.label for r in relevant[0].rows
                      if r.role in ("data", "preloaded")]
            if stmt == "IS":
                target_lines_pnl = labels
            else:
                target_lines_bs = labels
        st.caption(f"✅ Dropdowns: {len(target_lines_pnl)} P&L lines, {len(target_lines_bs)} BS lines.")
    except Exception as e:
        st.warning(f"Could not read target template: {e}")
else:
    st.warning("⚠️ No target template uploaded — upload on **📂 Upload Files**.")

all_targets = sorted(set([""] + target_lines_pnl + target_lines_bs))


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
            options=all_targets if all_targets else [""],
            width="medium", required=False),
    },
    height=600,
    key="mapping_editor_main",
)

# ── Save ─────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 3])
with col1:
    save_btn = st.button("💾 Save overrides", type="primary", use_container_width=True)
with col2:
    if active_profile:
        st.caption(f"Saves per-company to profile **{active_profile['name']}** in MongoDB.")
    else:
        st.caption("Session only — pick a profile on **🗂️ Profiles** to persist.")

if save_btn:
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

st.divider()
if template_loaded:
    st.info("👉 Next: **💾 Generate Linked Workbook**.")
else:
    st.warning("⚠️ No target template uploaded. Go to **📂 Upload Files**.")
