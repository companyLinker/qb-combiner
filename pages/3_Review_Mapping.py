"""Page 3: Review & Override Mappings — per-company view.

Clicking "Save Overrides" ALWAYS shows a profile-chooser dialog
so the user explicitly picks (or creates) a profile every time.
"""

import io
import hashlib
import streamlit as st
import openpyxl
import pandas as pd
from lib.mapping_rules import map_pnl, map_bs
from lib.template_discovery import discover_template
from lib import profiles as P
from lib import db as dblib
from lib.ui import hide_streamlit_elements


@st.cache_data(show_spinner=False, max_entries=4)
def _cached_template_lines_p3(target_bytes_hash: str, _target_bytes: bytes):
    """Discover P&L and BS target lines from template. Cached by content hash."""
    pnl_lines: list = []
    bs_lines: list  = []
    try:
        wb_tpl = openpyxl.load_workbook(io.BytesIO(_target_bytes), data_only=False)
        year_sheets = discover_template(wb_tpl)
        for stmt_key, is_pnl in [("IS", True), ("BS", False)]:
            relevant = sorted([s for s in year_sheets if s.statement == stmt_key],
                              key=lambda s: -(s.year or 0))
            if not relevant:
                continue
            labels = [
                r.label for r in relevant[0].rows
                if r.role in ("data", "preloaded", "subtotal") and r.label.strip()
            ]
            if is_pnl:
                pnl_lines = labels
            else:
                bs_lines = labels
    except Exception:
        pass
    return pnl_lines, bs_lines


hide_streamlit_elements()

st.title("🔗 Step 3 — Review & Override Mappings")

if "qb_data" not in st.session_state:
    st.warning("👈 Upload files first on **📂 Upload Files**.")
    st.stop()

# ── DB status ────────────────────────────────────────────────────────────────
db_online = dblib.is_connected()
if not db_online:
    st.warning(f"MongoDB unavailable — session-only mode. ({dblib.connection_error()})")

# ── Load active profile (for reading existing mappings into table) ────────────
if db_online and not st.session_state.get("active_profile_id"):
    _first = P.list_profiles()
    if _first:
        st.session_state.active_profile_id = str(_first[0]["_id"])

active_profile = None
saved_lookup: dict        = {}
entity_saved_lookup: dict = {}
if db_online:
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        active_profile      = P.get_profile(active_id)
        saved_lookup        = P.mapping_lookup(active_id)
        entity_saved_lookup = P.entity_mapping_lookup(active_id)
        
        # Load row pivot overrides from database if switched profile
        if st.session_state.get("loaded_row_overrides_profile_id") != active_id:
            db_ro = {}
            for k, v in saved_lookup.items():
                if k.startswith("__template_row_override__|"):
                    parts = k.split("|", 2)
                    if len(parts) == 3:
                        db_ro[f"{parts[1]}|{parts[2]}"] = v
            st.session_state.row_pivot_overrides = db_ro
            st.session_state.loaded_row_overrides_profile_id = active_id

if active_profile:
    st.success(
        f"📂 Loaded from profile: **{active_profile['name']}** — "
        f"{len(saved_lookup)} generic + {len(entity_saved_lookup)} entity-specific mappings."
    )
else:
    st.info("No active profile — session-only. Mappings will be loaded from session.")

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
        _tb = st.session_state.target_bytes
        _tb_hash = hashlib.md5(_tb).hexdigest()
        target_lines_pnl, target_lines_bs = _cached_template_lines_p3(_tb_hash, _tb)
        st.caption(
            f"✅ Template dropdowns: {len(target_lines_pnl)} P&L lines, "
            f"{len(target_lines_bs)} BS lines loaded."
        )
    except Exception as e:
        st.warning(f"Could not read target template: {e}")
else:
    st.warning("⚠️ No target template uploaded — upload on **📂 Upload Files**.") 

if target_lines_pnl or target_lines_bs:
    all_targets = sorted(set([""] + target_lines_pnl + target_lines_bs))
else:
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

# ── Save — dialog appears every time ──────────────────────────────────────────
st.divider()

if "p3_save_dialog_open" not in st.session_state:
    st.session_state.p3_save_dialog_open = False

save_col, hint_col = st.columns([1, 3])
with save_col:
    if st.button("💾 Save Overrides", type="primary",
                 use_container_width=True, key="p3_save_btn"):
        st.session_state.p3_save_dialog_open = True
with hint_col:
    if db_online:
        st.caption("You will be asked **which profile** to save to every time.")
    else:
        st.warning("MongoDB offline — mappings saved to session only.")

# ── Profile chooser dialog ─────────────────────────────────────────────────────
if st.session_state.get("p3_save_dialog_open"):
    st.markdown("---")
    with st.container(border=True):
        st.markdown("#### 📂 Save overrides to which profile?")

        if not db_online:
            st.warning("MongoDB is offline. Mappings will be saved to this browser session only.")
            oc1, oc2 = st.columns([1, 1])
            with oc1:
                if st.button("✅ Save to session", type="primary",
                             use_container_width=True, key="p3_sess_confirm"):
                    new_ov = dict(entity_session_overrides)
                    n = 0
                    for _, r in edited.iterrows():
                        ekey   = r["entity_key"]
                        chosen = (r["Target Line"] or "").strip()
                        if chosen:
                            new_ov[ekey] = chosen
                            n += 1
                        else:
                            new_ov.pop(ekey, None)
                    st.session_state.entity_mapping_overrides = new_ov
                    st.session_state.p3_save_dialog_open = False
                    st.success(f"✅ Saved {n} overrides to session.")
                    st.rerun()
            with oc2:
                if st.button("❌ Cancel", use_container_width=True, key="p3_sess_cancel"):
                    st.session_state.p3_save_dialog_open = False
                    st.rerun()
        else:
            _profiles_now = P.list_profiles()
            CREATE_NEW    = "__CREATE_NEW__"
            _pids   = [str(p["_id"]) for p in _profiles_now] + [CREATE_NEW]
            _plbls  = [p["name"]     for p in _profiles_now] + ["➕ Create new profile…"]

            _cur = st.session_state.get("active_profile_id", "")
            _def = _pids.index(_cur) if _cur in _pids else 0

            sel_id = st.selectbox(
                "Select profile to save into:",
                options=_pids,
                format_func=lambda x: _plbls[_pids.index(x)],
                index=_def,
                key="p3_dialog_sel",
            )

            new_name = ""
            if sel_id == CREATE_NEW:
                new_name = st.text_input(
                    "New profile name",
                    placeholder="Enter a name for the new profile…",
                    key="p3_dialog_new_name",
                )

            bc1, bc2 = st.columns([1, 1])
            with bc1:
                confirm = st.button("✅ Confirm Save", type="primary",
                                    use_container_width=True, key="p3_confirm")
            with bc2:
                if st.button("❌ Cancel", use_container_width=True, key="p3_cancel"):
                    st.session_state.p3_save_dialog_open = False
                    st.rerun()

            if confirm:
                if sel_id == CREATE_NEW:
                    if not new_name.strip():
                        st.warning("⚠️ Enter a name for the new profile first.")
                        st.stop()
                    tid   = str(P.create_profile(new_name.strip()))
                    tname = new_name.strip()
                else:
                    tid   = sel_id
                    tname = _plbls[_pids.index(sel_id)]

                st.session_state.active_profile_id = tid

                new_ov = dict(entity_session_overrides)
                n = 0
                for _, r in edited.iterrows():
                    ekey   = r["entity_key"]
                    chosen = (r["Target Line"] or "").strip()
                    parts  = ekey.split("|", 4)
                    ename  = parts[2] if len(parts) == 5 else ""
                    stmt   = parts[1] if len(parts) >= 2 else ""
                    bc_v   = parts[3] if len(parts) >= 4 else ""
                    lbl    = parts[4] if len(parts) >= 5 else ""

                    if chosen:
                        new_ov[ekey] = chosen
                        n += 1
                        if ename:
                            P.upsert_entity_mapping(
                                profile_id=tid, entity=ename, statement=stmt,
                                breadcrumb=bc_v, qb_account=lbl,
                                target_line=chosen, source="manual",
                            )
                    else:
                        new_ov.pop(ekey, None)
                        if ename:
                            try:
                                P.delete_entity_mapping(tid, ename, stmt, bc_v, lbl)
                            except Exception:
                                pass

                st.session_state.entity_mapping_overrides = new_ov
                
                # Save row-level pivot overrides
                if db_online:
                    d = dblib.get_db()
                    if d is not None:
                        d.mappings.delete_many({
                            "profile_id": P.ObjectId(tid),
                            "statement": "__template_row_override__"
                        })
                        for rk, val in st.session_state.get("row_pivot_overrides", {}).items():
                            if val:
                                parts = rk.split("|", 1)
                                if len(parts) == 2:
                                    sheet_name, row_idx = parts[0], parts[1]
                                    P.upsert_mapping(
                                        profile_id=tid,
                                        statement="__template_row_override__",
                                        breadcrumb=sheet_name,
                                        qb_account=row_idx,
                                        target_line=val,
                                        source="manual",
                                    )
                        P._invalidate_mapping_cache(tid)
                st.session_state.loaded_row_overrides_profile_id = tid

                st.session_state.p3_save_dialog_open = False
                st.success(f"✅ Saved {n} overrides and template row formula overrides to profile **{tname}**.")
                st.rerun()

if template_loaded:
    st.info("👉 Next: **💾 Generate Linked Workbook**.")
else:
    st.warning("⚠️ No target template uploaded. Go to **📂 Upload Files**.")
