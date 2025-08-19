import streamlit as st
import pandas as pd
import re
from fuzzywuzzy import fuzz
import streamlit.components.v1 as components

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

# Use columns to place buttons side by side
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
# No need for the third column as it's empty in the screenshot

st.markdown("---")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error loading the Excel file: {e}. Please ensure it is a valid .xlsx file.")
        st.stop()
    
    # Collect component types for dropdown
    component_types = set()
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")
            df.columns = df.columns.str.strip()
            if "Component Type" in df.columns:
                component_types.update(df["Component Type"].dropna().unique())
        except Exception as e:
            st.warning(f"Skipping sheet '{sheet}' due to an error: {e}")
    component_types = sorted(list(component_types))

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
    selected_sheets = st.multiselect("Select sheets to search", options=xls.sheet_names, default=xls.sheet_names)

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
            if label == "Others" and not term:
                continue
            if term:
                st.header(f"🔍 Results for {label}: {term}")
                results = {}
                for sheet_name in selected_sheets:
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
                    except Exception as e:
                        st.warning(f"Skipping sheet '{sheet_name}' due to an error: {e}")
                        continue
                        
                    first_row = df.iloc[0].astype(str).str.lower().tolist()
                    if any("general search" in cell for cell in first_row):
                        df = df[1:]
                    site_headers = df.columns.tolist()[1:]
                    site_results = {}
                    for site in site_headers:
                        if site in df.columns:
                            matched_rows = df[df[site].astype(str).apply(
                                lambda x: fuzz.partial_ratio(term.lower(), str(x).lower()) >= fuzzy_threshold
                            )].copy()
                            if label == "Others" and selected_component_type:
                                if "Component Type" in matched_rows.columns:
                                    matched_rows = matched_rows[matched_rows["Component Type"] == selected_component_type]
                            if label == "Others" and quantity_filter > 0 and "Quantity" in matched_rows.columns:
                                matched_rows = matched_rows[matched_rows["Quantity"] >= quantity_filter]
                            if label == "Others" and status_filter and "Hardware Status" in matched_rows.columns:
                                matched_rows = matched_rows[
                                    matched_rows["Hardware Status"].astype(str).str.contains(status_filter, case=False)
                                ]
                            
                            if "Link" not in matched_rows.columns:
                                matched_rows["Link"] = ""
                            matched_rows["Link"] = matched_rows[site].apply(extract_link)
                            
                            for col in ["Component Type", "Quantity", "Hardware Status", "Location", "Notes"]:
                                if col not in matched_rows.columns:
                                    matched_rows[col] = ""
                            site_results[site] = matched_rows[[site, "Component Type", "Quantity", "Hardware Status", "Location", "Notes", "Link"]]
                            for _, row in matched_rows.iterrows():
                                export_data.append({
                                    "Sheet": sheet_name,
                                    "Search Category": label,
                                    "Site": site,
                                    "Component": row[site],
                                    "Component Type": row["Component Type"],
                                    "Quantity": row["Quantity"],
                                    "Hardware Status": row["Hardware Status"],
                                    "Location": row["Location"],
                                    "Notes": row["Notes"],
                                    "Link": row["Link"]
                                })
                    results[sheet_name] = site_results

                for sheet, sites in results.items():
                    st.subheader(f"📄 Sheet: {sheet}")
                    for site, data in sites.items():
                        if not data.empty:
                            color_class = ""
                            if "osa21" in site.lower():
                                color_class = "osa21-color"
                            elif "osa22" in site.lower():
                                color_class = "osa22-color"
                            elif "osa23" in site.lower():
                                color_class = "osa23-color"
                            
                            st.markdown(f'### <span class="site-header {color_class}">📍 Site: {site}</span>', unsafe_allow_html=True)
                            
                            with st.expander("Expand/Minimize Results"):
                                st.dataframe(
                                    data,
                                    column_config={
                                        data.columns[0]: st.column_config.TextColumn(
                                            label="Component",
                                            width="large",
                                        ),
                                        "Component Type": st.column_config.TextColumn(
                                            label="Component Type",
                                            width="medium",
                                        ),
                                        "Quantity": st.column_config.NumberColumn(
                                            label="Quantity",
                                            width="small",
                                        ),
                                        "Hardware Status": st.column_config.TextColumn(
                                            label="Status",
                                            width="small",
                                        ),
                                        "Location": st.column_config.TextColumn(
                                            label="Location",
                                            width="small",
                                        ),
                                        "Notes": st.column_config.TextColumn(
                                            label="Notes",
                                            width="medium",
                                        ),
                                        "Link": st.column_config.LinkColumn(
                                            label="Link",
                                            help="Click to open the link",
                                            width="large",
                                        )
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