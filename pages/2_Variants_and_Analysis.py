"""Page 2: Variants & Analysis — per-company mapping editor.

Each row in the table = one company × one QB account.
Company Name column added before Statement.
Target Line can be set per-company; saves to entity_mapping_overrides
(MongoDB entity_mappings collection when a profile is active).
"""

import io
import streamlit as st
import openpyxl
import pandas as pd

from lib.master_builder import build_master_workbook, build_variants_digest, copy_sheet_into_workbook
from lib.linked_builder import build_linked_workbook
from lib.mapping_rules import map_pnl, map_bs
from lib.template_discovery import discover_template
from lib import db as dblib
from lib import profiles as P


if st.secrets.get("APP_PASSWORD", None) and not st.session_state.get("authed"):
    st.warning("Please sign in from the home page first.")
    st.stop()

st.title("📊 Step 2 — Variants & Analysis")

if "qb_data" not in st.session_state:
    st.warning("Upload files first on **📂 Upload Files**.")
    st.stop()

qb_data = st.session_state.qb_data

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
        active_profile       = P.get_profile(active_id)
        saved_lookup         = P.mapping_lookup(active_id)
        entity_saved_lookup  = P.entity_mapping_lookup(active_id)

# Merge: MongoDB base + session overrides on top
session_overrides: dict        = dict(st.session_state.get("mapping_overrides", {}))
entity_session_overrides: dict = dict(st.session_state.get("entity_mapping_overrides", {}))
generic_lookup = {**saved_lookup, **session_overrides}
entity_lookup  = {**entity_saved_lookup, **entity_session_overrides}

# ── Build master + digest ────────────────────────────────────────────────────
with st.spinner("Building consolidated master + variants digest..."):
    master_buf, pnl_pivot, bs_pivot = build_master_workbook(
        qb_data,
        mapping_overrides=generic_lookup,
        entity_mapping_overrides=entity_lookup,
    )
    digest_buf = build_variants_digest(qb_data, pnl_pivot, bs_pivot)
st.session_state.pnl_pivot = pnl_pivot
st.session_state.bs_pivot  = bs_pivot

# ── Optionally append selected template sheets (with SUMIFS) to master wb ────
_target_bytes      = st.session_state.get("target_bytes")
_selected_sheets   = st.session_state.get("selected_template_sheets")
_entity_col_map    = st.session_state.get("template_entity_mapping")
_template_sheets_appended = False

if _target_bytes and _selected_sheets:
    try:
        with st.spinner("Appending template sheets to master workbook…"):
            linked_buf, _, _, _ = build_linked_workbook(
                qb_data,
                _target_bytes,
                mapping_overrides=generic_lookup,
                entity_mapping_overrides=entity_lookup,
                selected_sheets=_selected_sheets,
                entity_col_mapping=_entity_col_map or None,
            )
            # Load both workbooks and copy the selected sheets
            master_wb = openpyxl.load_workbook(io.BytesIO(master_buf.getvalue()))
            linked_wb = openpyxl.load_workbook(
                io.BytesIO(linked_buf.getvalue()), data_only=False
            )
            for sn in _selected_sheets:
                if sn in linked_wb.sheetnames:
                    copy_sheet_into_workbook(linked_wb, sn, master_wb)

            new_buf = io.BytesIO()
            master_wb.save(new_buf)
            new_buf.seek(0)
            master_buf = new_buf
            _template_sheets_appended = True
    except Exception as _err:
        st.warning(f"⚠️ Could not append template sheets to master workbook: {_err}")

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
    except Exception as e:
        st.warning(f"Could not auto-discover target template: {e}")

# ── Metrics ──────────────────────────────────────────────────────────────────
def leaf_index(pivot):
    return {k: v for k, v in pivot.items()
            if not v["is_section"] and not v["is_total"]}

pnl_leaves = leaf_index(pnl_pivot)
bs_leaves  = leaf_index(bs_pivot)
variants   = {info.get("variant", "single") for info in qb_data.values()}
variant_label = "Dual-year CY+PY" if "triple" in variants else "Single-year"

c1, c2, c3, c4 = st.columns(4)
c1.metric("Entities", len(qb_data))
c2.metric("P&L accounts", len(pnl_leaves))
c3.metric("BS accounts", len(bs_leaves))
total_abs = (sum(abs(sum(v["amounts"].values())) for v in pnl_leaves.values()) +
             sum(abs(sum(v["amounts"].values())) for v in bs_leaves.values()))
c4.metric("Total $ flowing", f"${total_abs:,.0f}")
st.caption(f"QB format: **{variant_label}**")
if active_profile:
    st.caption(f"📂 Profile: **{active_profile['name']}** — "
               f"{len(saved_lookup)} generic + {len(entity_saved_lookup)} entity-specific mappings loaded.")

st.divider()
st.markdown("### Downloads")

_master_label = (
    f"⬇ Master Consolidated Workbook (+{len(_selected_sheets)} template sheets)"
    if _template_sheets_appended
    else "⬇ Master Consolidated Workbook"
)
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        _master_label,
        master_buf.getvalue(),
        "00_MASTER_Consolidated.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    if _template_sheets_appended:
        st.caption(
            f"✅ Includes template tabs: **{', '.join(_selected_sheets)}** with SUMIFS data."
        )
    elif _target_bytes and _selected_sheets is None:
        st.caption("ℹ️ Configure template sheets in **📂 Upload Files** to include them here.")
with col2:
    st.download_button(
        "⬇ Variants Digest",
        digest_buf.getvalue(),
        "01_CoA_Variants_Digest.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ── Editable mapping table — per company ─────────────────────────────────────
st.divider()
st.markdown("### Map accounts to target template lines")
st.caption("Each row = one **Company × QB Account**. Set a Target Line per company. "
           "Sorted by largest $ first.")

if template_loaded:
    st.caption(f"Dropdowns from template: {len(target_lines_pnl)} P&L lines, "
               f"{len(target_lines_bs)} BS lines.")
else:
    st.warning("⚠️ Upload a target template on **📂 Upload Files** to enable dropdowns.")


def make_rows():
    rows = []
    for entity, info in qb_data.items():
        for stmt_kind, row_key, leaves in [
            ("P&L", "pnl_rows", pnl_leaves),
            ("BS",  "bs_rows",  bs_leaves),
        ]:
            for r in info.get(row_key, []):
                if r["is_section_only"] or r["is_total"]:
                    continue
                bc, lbl = r["breadcrumb"], r["label"]

                if stmt_kind == "P&L":
                    auto, conf = map_pnl(bc, lbl)
                    if auto == "__SKIP__":
                        continue
                else:
                    auto, conf = map_bs(bc, lbl)

                entity_key  = f"E|{stmt_kind}|{entity}|{bc}|{lbl}"
                generic_key = f"{stmt_kind}|{bc}|{lbl}"

                # Lookup priority: entity-specific > generic > auto
                if entity_key in entity_lookup and entity_lookup[entity_key]:
                    effective  = entity_lookup[entity_key]
                    confidence = "entity-saved"
                elif generic_key in generic_lookup and generic_lookup[generic_key]:
                    effective  = generic_lookup[generic_key]
                    confidence = "saved"
                else:
                    effective  = auto or ""
                    confidence = conf

                amt = (r.get("amount_cy") or r.get("amount") or 0)

                rows.append({
                    "entity_key":    entity_key,
                    "generic_key":   generic_key,
                    "Company Name":  entity,
                    "Statement":     stmt_kind,
                    "QB Account":    lbl,
                    "Parent path":   bc,
                    "Total $":       amt,
                    "abs_total":     abs(amt),
                    "Confidence":    confidence,
                    "Auto Suggestion": auto or "",
                    "Target Line":   effective,
                })
    rows.sort(key=lambda r: -r["abs_total"])
    return rows


rows = make_rows()
df   = pd.DataFrame(rows)

# Filters
chip_c1, chip_c2, chip_c3, chip_c4 = st.columns([2, 2, 2, 4])
with chip_c1:
    filter_stmt = st.radio("Statement", ["All", "P&L", "BS"],
                           horizontal=True, label_visibility="collapsed")
with chip_c2:
    filter_conf = st.radio("Confidence", ["All", "Need review", "Saved"],
                           horizontal=True, label_visibility="collapsed")
with chip_c3:
    companies = ["All"] + sorted(qb_data.keys())
    filter_co = st.selectbox("Company", companies, label_visibility="collapsed")
with chip_c4:
    search = st.text_input("Search", "", label_visibility="collapsed",
                            placeholder="🔍 Search account name...")

view = df.copy()
if filter_stmt == "P&L":
    view = view[view["Statement"] == "P&L"]
elif filter_stmt == "BS":
    view = view[view["Statement"] == "BS"]
if filter_conf == "Need review":
    view = view[view["Confidence"] == "REVIEW"]
elif filter_conf == "Saved":
    view = view[view["Confidence"].isin(["saved", "entity-saved", "manual"])]
if filter_co != "All":
    view = view[view["Company Name"] == filter_co]
if search.strip():
    s = search.strip().lower()
    view = view[view["QB Account"].str.lower().str.contains(s, na=False) |
                view["Parent path"].str.lower().str.contains(s, na=False)]

view = view.drop(columns=["abs_total"], errors="ignore")
all_targets = sorted(set([""] + target_lines_pnl + target_lines_bs))

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
        "Parent path":    st.column_config.TextColumn(disabled=True, width="large"),
        "Total $":        st.column_config.NumberColumn(disabled=True, format="$%.0f", width="small"),
        "Confidence":     st.column_config.TextColumn(disabled=True, width="small"),
        "Auto Suggestion":st.column_config.TextColumn(disabled=True, width="medium"),
        "Target Line":    st.column_config.SelectboxColumn(
            "Target Line",
            options=all_targets if all_targets else [""],
            width="medium", required=False),
    },
    height=600,
    key="variants_editor_main",
)

# ── Save ─────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 3])
with col1:
    save_btn = st.button("💾 Save mappings", type="primary", use_container_width=True)
with col2:
    if active_profile:
        st.caption(f"Saves per-company mappings to profile **{active_profile['name']}** in MongoDB.")
    elif dblib.is_connected():
        st.caption("No profile active — session only. Go to **🗂️ Profiles** to persist.")
    else:
        st.warning("MongoDB offline — session-only mode.")

if save_btn:
    new_entity_overrides  = dict(entity_session_overrides)
    n_entity_saved = 0

    for _, r in edited.iterrows():
        ekey    = r["entity_key"]
        gkey    = r["generic_key"]
        chosen  = (r["Target Line"] or "").strip()
        # Parse entity_key: "E|stmt|entity|bc|lbl"
        parts = ekey.split("|", 4)  # ["E", stmt, entity, bc, lbl]
        entity_name = parts[2] if len(parts) == 5 else ""
        stmt        = parts[1] if len(parts) >= 2 else ""
        bc          = parts[3] if len(parts) >= 4 else ""
        lbl         = parts[4] if len(parts) >= 5 else ""

        if chosen:
            new_entity_overrides[ekey] = chosen
            n_entity_saved += 1
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
    if active_profile and n_entity_saved:
        st.success(f"✅ Saved {n_entity_saved} per-company mappings to profile **{active_profile['name']}**.")
    else:
        st.success(f"✅ Saved {len(new_entity_overrides)} entity overrides in session.")

st.divider()
if template_loaded:
    st.info("👉 Next: **💾 Generate Linked Workbook**.")
else:
    st.warning("Upload a target template on **📂 Upload Files** before generating.")
