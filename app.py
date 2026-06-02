"""QB Combiner — Streamlit app entry point.

Run locally:
    streamlit run app.py

Or with Docker:
    docker compose up
"""

import streamlit as st
from lib import db as dblib
from lib import profiles as P
from lib.ui import hide_streamlit_elements


# ------------------------- Page config -------------------------
st.set_page_config(
    page_title="QuickBooks Combiner",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
hide_streamlit_elements()


# ------------------------- Password gate -------------------------
def check_password():
    expected = st.secrets.get("APP_PASSWORD", None) if hasattr(st, "secrets") else None
    if not expected:
        return True
    if st.session_state.get("authed"):
        return True

    st.markdown("## 🔒 Sign in")
    st.markdown("Enter the team password to use the QB Combiner.")
    pw = st.text_input("Password", type="password", key="pw_input")
    if st.button("Sign in", type="primary"):
        if pw == expected:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


if not check_password():
    st.stop()


# ------------------------- Landing -------------------------
st.title("📊 QuickBooks → Combination Workbook")
st.markdown(
    """
A pipeline to consolidate per-entity QuickBooks P&L and Balance Sheet exports
into a single combination workbook with **SUMIFS-linked formulas** to your
tax/financial template.

### How it works

1. **🗂️ Profiles** *(optional but recommended)* — create or pick a named profile so your mapping decisions persist across years.
2. **📂 Upload Files** — drop your QuickBooks exports + target combination template.
3. **📊 Variants & Analysis** — see the consolidated data and chart-of-accounts variants.
4. **🔗 Review Mapping** — auto-mapping handles ~99% of P&L and ~93% of BS; the rest you confirm by hand.
5. **💾 Generate** — produce the linked workbook and download it.

Use the **sidebar** to navigate.
"""
)


st.divider()
st.markdown("### 📋 Session status")

active_profile = None
mongo_ok = dblib.is_connected()
if mongo_ok:
    active_id = st.session_state.get("active_profile_id")
    if active_id:
        try:
            active_profile = P.get_profile(active_id)
        except Exception:
            active_profile = None

qb_loaded = "qb_data" in st.session_state and st.session_state.qb_data
target_loaded = "target_bytes" in st.session_state and st.session_state.target_bytes
overrides_count = len(st.session_state.get("mapping_overrides", {}))

cols = st.columns(4)
with cols[0]:
    if active_profile:
        n_maps = P.mapping_count(str(active_profile["_id"]))
        st.success(f"✅ Profile: **{active_profile['name']}** ({n_maps} saved)")
    elif mongo_ok:
        st.info("⏳ No profile selected — pick one on 🗂️ Profiles")
    else:
        st.warning("⚠️ MongoDB offline — session-only mode")
with cols[1]:
    if qb_loaded:
        st.success(f"✅ {len(st.session_state.qb_data)} QB files loaded")
    else:
        st.info("⏳ No QB files yet")
with cols[2]:
    if target_loaded:
        st.success("✅ Target template loaded")
    else:
        st.info("⏳ No target template yet")
with cols[3]:
    if overrides_count:
        st.success(f"✅ {overrides_count} session overrides")
    else:
        st.info("⏳ No session overrides")


if not mongo_ok:
    with st.expander("ℹ️ Why is MongoDB offline?"):
        st.markdown(
            f"Reason: `{dblib.connection_error()}`\n\n"
            "Profiles and mapping persistence require MongoDB. The app still works without it — "
            "just in session-only mode (mappings forgotten when you close the tab). "
            "To enable persistence, run `docker compose up` (which starts MongoDB automatically) "
            "or set `MONGODB_URI` in `.streamlit/secrets.toml`. See `DEPLOYMENT.md`."
        )


if qb_loaded:
    if st.button("🔄 Reset session"):
        for key in list(st.session_state.keys()):
            if key not in ("authed", "active_profile_id"):
                del st.session_state[key]
        st.rerun()
