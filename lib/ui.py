import streamlit as st

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
