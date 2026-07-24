import hashlib

import pandas as pd
import streamlit as st


# The 4 linear steps of the main wizard flow (Profiles and Template Inspector
# are non-linear utility pages and don't appear here).
STEPS = [
    (1, "📂", "Upload Files", "pages/1_Upload_Files.py"),
    (2, "📊", "Variants & Analysis", "pages/2_Variants_and_Analysis.py"),
    (3, "🔗", "Review Mapping", "pages/3_Review_Mapping.py"),
    (4, "💾", "Generate Workbook", "pages/4_Generate_Linked_Workbook.py"),
]


def render_step_header(current: int):
    """Show all 4 wizard steps at once: done steps in solid green, the
    current step boxed, upcoming steps muted. Pure native components (no
    custom CSS) so it can't fight Streamlit's own DOM across versions."""
    cols = st.columns(len(STEPS))
    for col, (num, icon, label, _path) in zip(cols, STEPS):
        with col:
            if num < current:
                st.success(f"✅ {num}. {label}")
            elif num == current:
                with st.container(border=True):
                    st.markdown(f"**{icon} {num}. {label}**")
            else:
                st.caption(f"{num}. {label}")


def render_next_step(ready: bool, target_page: str, label: str, not_ready_msg: str):
    """A strong, high-contrast 'you're good to go' cue plus an actual
    navigation button when ready — instead of a permanently pale st.info
    hint that looks the same whether or not the user has anything to do."""
    st.divider()
    if ready:
        st.success("✅ This step looks good.")
        if st.button(f"➡️ Continue to {label}", type="primary", width="stretch"):
            st.switch_page(target_page)
    else:
        st.info(not_ready_msg)


def rows_hash(df: pd.DataFrame, columns: list) -> str:
    """Fast vectorized row hash (pandas' own C-accelerated hasher) over just
    the given columns — used to detect whether an editable table has any
    unsaved changes."""
    return hashlib.md5(
        pd.util.hash_pandas_object(df[columns], index=False).values.tobytes()
    ).hexdigest()


def is_dirty(current_df: pd.DataFrame, baseline_df: pd.DataFrame, columns: list) -> bool:
    """True if current_df differs from baseline_df over the given columns.
    Pass the pre-edit dataframe handed to st.data_editor(...) as the
    baseline — it's rebuilt fresh from saved state every run, so it already
    represents "last saved," no separate snapshot needs to be tracked."""
    return rows_hash(current_df, columns) != rows_hash(baseline_df, columns)


def hide_streamlit_elements():
    """Injects custom CSS to hide the Streamlit main menu, footer, deployment/fork buttons, 
    and GitHub toolbar for a clean, professional application look.
    
    IMPORTANT: We do NOT hide the full `header` element because that also hides
    the sidebar collapse/expand toggle button on Streamlit Cloud, making the
    sidebar permanently inaccessible. Instead we hide only specific children.
    """
    st.markdown(
        """
        <style>
        /* Hide main menu and footer */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        /* Hide the stAppHeader bar itself but keep the sidebar toggle button visible.
           We target specific children of the header, not the whole header block. */
        [data-testid="stHeader"] > div:not([data-testid="stSidebarNav"]):not([data-testid="collapsedControl"]) {
            display: none !important;
        }
        /* Hide the top decoration bar (the thin colored line at very top) */
        [data-testid="stDecoration"] {display: none !important;}
        /* Hide the top toolbar (hamburger + three-dot menu area) */
        [data-testid="stToolbar"] {visibility: hidden !important; height: 0 !important;}

        /* Hide the Streamlit Cloud action buttons/ribbons */
        .stDeployButton {display: none !important;}
        .stViewerBadge {display: none !important;}

        /* Force hiding of the "Manage app" iframe button */
        iframe[title="Manage app"] {display: none !important;}

        /* Hide the bottom Streamlit Cloud creator avatar/profile */
        div[class*="_profileContainer_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        div[class*="_profilePreview_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        img[class*="_profileImage_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        img[data-testid="appCreatorAvatar"] {display: none !important;}
        a[href*="share.streamlit.io"] {display: none !important;}
        div.stApp [class*="_profile"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important;}

        /* Ensure sidebar collapse button is always visible and clickable */
        [data-testid="collapsedControl"] {
            visibility: visible !important;
            display: flex !important;
            opacity: 1 !important;
            z-index: 9999 !important;
        }
        button[kind="header"] {
            visibility: visible !important;
            display: flex !important;
            opacity: 1 !important;
        }

        /* Adjust layout to remove top padding from header hiding */
        .block-container {
            padding-top: 2rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
