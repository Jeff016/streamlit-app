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
        xls = pd.ExcelFile(uploaded_file_obj, engine="openpyxl")
        
        sheet_names = xls.sheet_names
        df_cache = {}
        component_types = set()

        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df_cache[sheet] = df.copy() 
            df.columns = df.columns.str.strip()
            if "Component Type" in df.columns:
                component_types.update(df["Component Type"].dropna().unique())
        
        return True, df_cache, sorted(list(component_types)), sheet_names, None
    except Exception as e:
        error_message = f"Error loading the Excel file: {e}. Please ensure it is a valid, uncorrupted .xlsx file."
        return False, {}, [], [], error_message

# Process file only if a new file is uploaded
if uploaded_file is not None:
    current_file_id = uploaded_file.file_id
    if st.session_state.last_uploaded_file_id != current_file_id:
        with st.spinner('Processing your Excel file... This may take a moment.'):
            (
                success,
                df_cache,
                component_types,
                sheet_names,
                error,
            ) = load_data(uploaded_file)

            if success:
                st.session_state.data_loaded = True
                st.session_state.df_cache = df_cache
                st.session_state.component_types = component_types
                st.session_state.xls_sheet_names = sheet_names
                st.session_state.last_uploaded_file_id = current_file_id
                st.success("File uploaded and data loaded successfully!")
                st.rerun() 
            else:
                st.error(error)
                st.session_state.data_loaded = False
                st.session_state.last_uploaded_file_id = None

# Main application logic runs if data is successfully loaded
if st.session_state.data_loaded:
    component_types = st.session_state.component_types
    xls_sheet_names = st.session_state.xls_sheet_names

    # Search inputs
    processor_term = st.text_input("🔍 Processor")
    ram_term = st.text_input("🔍 RAM")
    hard_drive_term = st.text_input("🔍 Hard Drive")
    remote_mgmt_term = st.text_input("🔍 Remote Mgmt Card")
    drive_controller_term = st.text_input("🔍 Drive Controller")

    # Others section
    st.markdown("---")
    st.subheader("🔍 Others Search")
    selected_component_type = st.selectbox("Select Component Category", options=[""] + component_types)
    others_term = st.text_input("Search term for Others")
    quantity_filter = st.number_input("Minimum Quantity", min_value=0, value=0)
    status_filter = st.text_input("Hardware Status (optional)")
    fuzzy_threshold = st.slider("Fuzzy match threshold", min_value=70, max_value=100, value=85)
    selected_sheets = st.multiselect("Select sheets to search", options=xls_sheet_names, default=xls_sheet_names)

    search_terms = {
        "Processor": processor_term,
        "RAM": ram_term,
        "Hard Drive": hard_drive_term,
        "Remote Mgmt Card": remote_mgmt_term,
        "Drive Controller": drive_controller_term,
        "Others": others_term
    }
    
    export_data = []

    def extract_link(text):
        if isinstance(text, str):
            md_match = re.search(r'\[(.*?)\]\((https?://[^\s]+)\)', text)
            if md_match:
                return md_match.group(2)
            plain_match = re.search(r'(https?://[^\s]+)', text)
            if plain_match:
                return plain_match.group(1)
        return ""

    if st.button("🔎 Search Components"):
        for label, term in search_terms.items():
            current_term = term if label != "Others" else others_term
            if not current_term:
                continue
            
            st.header(f"🔍 Results for {label}: {current_term}")
            results = {}
            for sheet_name in selected_sheets:
                try:
                    df = st.session_state.df_cache[sheet_name].copy()
                except KeyError:
                    st.warning(f"Sheet '{sheet_name}' not found in cache. Skipping.")
                    continue
                
                first_row = df.iloc[0].astype(str).str.lower().tolist()
                if any("general search" in cell for cell in first_row):
                    df = df.iloc[1:].reset_index(drop=True)

                site_headers = [col for col in df.columns if col not in ["Component Type", "Quantity", "Hardware Status", "Location", "Notes"]]
                
                site_results = {}
                for site in site_headers:
                    if site in df.columns:
                        matched_rows = df[df[site].astype(str).apply(
                            lambda x: fuzz.partial_ratio(current_term.lower(), str(x).lower()) >= fuzzy_threshold
                        )].copy()
                        
                        if label == "Others":
                            if selected_component_type:
                                matched_rows = matched_rows[matched_rows["Component Type"] == selected_component_type]
                            if quantity_filter > 0 and "Quantity" in matched_rows.columns:
                                matched_rows["Quantity"] = pd.to_numeric(matched_rows["Quantity"], errors='coerce').fillna(0)
                                matched_rows = matched_rows[matched_rows["Quantity"] >= quantity_filter]
                            if status_filter and "Hardware Status" in matched_rows.columns:
                                matched_rows = matched_rows[
                                    matched_rows["Hardware Status"].astype(str).str.contains(status_filter, case=False, na=False)
                                ]
                        
                        if not matched_rows.empty:
                            matched_rows["Link"] = matched_rows[site].apply(extract_link)
                            
                            display_cols = [site, "Component Type", "Quantity", "Hardware Status", "Location", "Notes", "Link"]
                            final_cols = [col for col in display_cols if col in matched_rows.columns]
                            
                            site_results[site] = matched_rows[final_cols]
                            
                            for _, row in matched_rows.iterrows():
                                export_row = {
                                    "Sheet": sheet_name, "Search Category": label, "Site": site,
                                    "Component": row.get(site), "Component Type": row.get("Component Type"),
                                    "Quantity": row.get("Quantity"), "Hardware Status": row.get("Hardware Status"),
                                    "Location": row.get("Location"), "Notes": row.get("Notes"), "Link": row.get("Link")
                                }
                                export_data.append(export_row)

                results[sheet_name] = site_results

            for sheet, sites in results.items():
                st.subheader(f"📄 Sheet: {sheet}")
                for site, data in sites.items():
                    if not data.empty:
                        color_class = ""
                        if "osa21" in site.lower(): color_class = "osa21-color"
                        elif "osa22" in site.lower(): color_class = "osa22-color"
                        elif "osa23" in site.lower(): color_class = "osa23-color"
                        
                        st.markdown(f'### <span class="site-header {color_class}">📍 Site: {site}</span>', unsafe_allow_html=True)
                        
                        with st.expander("Expand/Minimize Results"):
                            st.dataframe(
                                data,
                                column_config={
                                    data.columns[0]: st.column_config.TextColumn(label="Component", width="large"),
                                    "Link": st.column_config.LinkColumn(label="Link", help="Click to open link", width="large")
                                },
                                hide_index=True,
                                use_container_width=True
                            )
                    else:
                        st.info(f"No matching components found for {site}.")

        if export_data:
            st.markdown("---")
            st.subheader("📤 Export Results")
            export_df = pd.DataFrame(export_data)
            csv = export_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV", data=csv, file_name="search_results.csv", mime="text/csv")

else:
    st.info("Please upload your Excel file to start.")
