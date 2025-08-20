import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
import streamlit.components.v1 as components
import io

st.set_page_config(layout="wide")

# Inject custom CSS for site headers and button layout
st.markdown("""
<style>
.site-header {
    font-size: 24px;
    font-weight: bold;
}
.osa21-color {
    color: #FF5733; /* Orange-Red for OSA21 */
}
.osa22-color {
    color: #33FF57; /* Green for OSA22 */
}
.osa23-color {
    color: #3385FF; /* Blue for OSA23 */
}
.stButton>button {
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

st.title("🔍 IBM Component Multi-Search Viewer")
st.markdown("Upload the Excel file and search for multiple components to view their associated links.")

# Add buttons with links
st.markdown("---")
st.subheader("Quick Links")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("##### INVENTORY SERVERS")
    if st.button("OSA21"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePKA', '_blank');
            </script>
            """,
            height=0
        )
    if st.button("OSA22"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePKu', '_blank');
            </script>
            """,
            height=0
        )
    if st.button("OSA23"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePKL', '_blank');
            </script>
            """,
            height=0
        )

with col2:
    st.markdown("##### FORMAT DRIVES")
    if st.button("OSA21 ", key="format_osa21"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePK9', '_blank');
            </script>
            """,
            height=0
        )
    if st.button("OSA22 ", key="format_osa22"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePKC', '_blank');
            </script>
            """,
            height=0
        )
    if st.button("OSA23 ", key="format_osa23"):
        components.html(
            """
            <script>
                window.open('https://ibm.biz/BdePKQ', '_blank');
            </script>
            """,
            height=0
        )

st.markdown("---")

# Use session state to handle the uploaded file and its data
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'df_cache' not in st.session_state:
    st.session_state.df_cache = {}
if 'component_types' not in st.session_state:
    st.session_state.component_types = []
if 'xls_sheet_names' not in st.session_state:
    st.session_state.xls_sheet_names = []
if 'last_uploaded_file_id' not in st.session_state:
    st.session_state.last_uploaded_file_id = None


# File uploader widget
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])


# Function to load and process the excel file
@st.cache_data(show_spinner=False)
def load_data(uploaded_file_obj):
    """Loads data from the uploaded Excel file and caches it."""
    try:
        # pd.ExcelFile can take the uploaded file object directly
        xls = pd.ExcelFile(uploaded_file_obj, engine="openpyxl")
        
        sheet_names = xls.sheet_names
        df_cache = {}
        component_types = set()

        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=
