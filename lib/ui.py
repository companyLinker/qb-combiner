import streamlit as st

def hide_streamlit_elements():
    """Injects custom CSS to hide the Streamlit main menu, footer, deployment/fork buttons, 
    and GitHub toolbar for a clean, professional application look.
    """
    st.markdown(
        """
        <style>
        /* Hide main menu, header, and footer */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .stAppHeader {visibility: hidden;}
        [data-testid="stHeader"] {display: none;}
        
        /* Hide the Streamlit Cloud action buttons/ribbons */
        .stDeployButton {display: none;}
        .stViewerBadge {display: none;}
        div[data-testid="stToolbar"] {display: none;}
        
        /* Force hiding of the top right menu and the "Manage app" button */
        iframe[title="Manage app"] {display: none !important;}
        div[data-testid="stDecoration"] {display: none !important;}
        
        /* Hide the bottom-left/bottom-right Streamlit Cloud creator avatar/profile */
        div[class*="_profileContainer_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        div[class*="_profilePreview_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        img[class*="_profileImage_"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important; visibility: hidden !important;}
        img[data-testid="appCreatorAvatar"] {display: none !important;}
        a[href*="share.streamlit.io"] {display: none !important;}
        div.stApp [class*="_profile"] {display: none !important; height: 0 !important; width: 0 !important; opacity: 0 !important;}
        
        /* Adjust layout to remove top padding from header hiding */
        .block-container {
            padding-top: 2rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
