"""Page 5: Mapping Profiles — create, select, rename, delete, export/import."""

import streamlit as st
from lib import db as dblib
from lib import profiles as P
from lib.ui import hide_streamlit_elements


hide_streamlit_elements()


st.title("🗂️ Mapping Profiles")
st.markdown(
    "A **profile** is a named bucket of mapping decisions (e.g., *AP NE 2025 Tax*, *Client B 2025*). "
    "Saved mappings persist across sessions — next year, the same QB accounts auto-fill from your prior decisions."
)


# --- Connection check ---
if not dblib.is_connected():
    st.error(
        f"**MongoDB is not connected** — profiles cannot be saved.\n\n"
        f"Reason: `{dblib.connection_error()}`\n\n"
        "Set `MONGODB_URI` in `.streamlit/secrets.toml` (e.g., `mongodb://mongo:27017`) "
        "and restart the app. See `DEPLOYMENT.md` for the docker-compose mongo setup."
    )
    st.stop()


# --- Create new profile ---
with st.expander("➕ Create a new profile", expanded=False):
    new_name = st.text_input("Profile name", placeholder="e.g., AP NE 2025 Tax")
    new_desc = st.text_input("Description (optional)", placeholder="e.g., 22 chicken LLCs, fiscal year 2025")
    if st.button("Create profile", type="primary", disabled=not new_name.strip()):
        existing = P.get_profile_by_name(new_name.strip())
        if existing:
            st.error(f"A profile named '{new_name.strip()}' already exists.")
        else:
            user_email = ""  # placeholder — wire up to real user when auth lands
            pid = P.create_profile(
                name=new_name.strip(),
                description=new_desc.strip(),
                target_template_name=st.session_state.get("target_filename", ""),
                created_by=user_email,
            )
            st.success(f"Created '{new_name.strip()}'.")
            st.session_state.active_profile_id = pid
            st.rerun()


# --- Active profile selector ---
profiles = P.list_profiles()
if not profiles:
    st.info("No profiles yet. Create one above to start saving mappings.")
    st.stop()

st.markdown("### Active profile")
active_id = st.session_state.get("active_profile_id")

# Build options list
labels = []
ids = []
for p in profiles:
    n_maps = P.mapping_count(str(p["_id"]))
    labels.append(f"{p['name']}  ({n_maps} saved mappings)")
    ids.append(str(p["_id"]))

current_idx = 0
if active_id and active_id in ids:
    current_idx = ids.index(active_id)

choice = st.radio("Choose the profile to use for the current session:",
                  options=range(len(labels)),
                  format_func=lambda i: labels[i],
                  index=current_idx)
chosen_id = ids[choice]

if chosen_id != active_id:
    st.session_state.active_profile_id = chosen_id
    P.touch_profile(chosen_id)
    st.success(f"Active profile: **{profiles[choice]['name']}**")


# --- Profile details ---
st.divider()
active = P.get_profile(chosen_id)
n_maps = P.mapping_count(chosen_id)
runs = P.list_runs(chosen_id, limit=5)

cols = st.columns(3)
cols[0].metric("Saved mappings", n_maps)
cols[1].metric("Generated runs", len(P.list_runs(chosen_id, limit=999)))
cols[2].metric("Template", active.get("target_template_name", "—") or "—")

if active.get("description"):
    st.caption(active["description"])


# --- Rename / Export / Delete ---
with st.expander("🔧 Manage this profile", expanded=False):
    new_label = st.text_input("Rename to", value=active["name"], key=f"rename_{chosen_id}")
    _name_unchanged = not new_label.strip() or new_label.strip() == active["name"]
    if st.button("Save name", disabled=_name_unchanged):
        P.rename_profile(chosen_id, new_label.strip())
        st.success("Renamed.")
        st.rerun()

    st.divider()
    json_blob = P.export_profile_json(chosen_id)
    st.download_button(
        "📤 Export profile as JSON",
        data=json_blob,
        file_name=f"profile_{active['name'].replace(' ', '_')}.json",
        mime="application/json",
    )

    st.divider()
    if st.checkbox(f"⚠️ I want to delete '{active['name']}' permanently", key=f"del_check_{chosen_id}"):
        if st.button("Delete profile", type="primary"):
            P.delete_profile(chosen_id)
            st.session_state.pop("active_profile_id", None)
            st.success("Profile deleted.")
            st.rerun()


# --- Import profile from JSON ---
with st.expander("📥 Import profile from JSON", expanded=False):
    up = st.file_uploader("Upload a previously-exported profile JSON", type=["json"])
    if up:
        try:
            new_id = P.import_profile_json(up.getvalue().decode("utf-8"))
            st.success(f"Imported profile (id={new_id}).")
            st.rerun()
        except Exception as e:
            st.error(f"Import failed: {e}")


# --- Recent runs ---
if runs:
    st.divider()
    st.markdown("### 📜 Recent runs")
    table = []
    for r in runs:
        table.append({
            "When": r["ran_at"].strftime("%Y-%m-%d %H:%M"),
            "By": r.get("ran_by", "—") or "—",
            "Entities": r["entities_count"],
            "Auto": r["auto_count"],
            "Manual": r["manual_count"],
            "REVIEW": r["review_count"],
            "Output": r.get("output_filename", "—"),
        })
    st.dataframe(table, width="stretch", hide_index=True)
