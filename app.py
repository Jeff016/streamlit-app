import streamlit as st
import pandas as pd
import re
from fuzzywuzzy import fuzz
import streamlit.components.v1 as components

st.set_page_config(layout="wide")

# Title and instructions
st.title("🔍 IBM Component Multi-Search Viewer")
st.markdown("Upload the Excel file and search for multiple components to view their associated links.")

# Quick Links section
st.markdown("---")
st.subheader("Quick Links")
col1, col2, _ = st.columns(3)

with col1:
    st.markdown("##### INVENTORY SERVERS")
    for label in ["OSA21", "OSA22", "OSA23"]:
        if st.button(label):
            components.html("", height=0)

with col2:
    st.markdown("##### FORMAT DRIVES")
    for label in ["OSA21", "OSA22", "OSA23"]:
        if st.button(label + " ", key=f"format_{label.lower()}"):
            components.html("", height=0)

st.markdown("---")

# File upload
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        st.success(f"Successfully loaded: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Error loading the Excel file: {e}. Please ensure it is a valid .xlsx file.")
        st.stop()

    # Collect component types
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
    search_terms = {
        "Processor": st.text_input("🔍 Processor"),
        "RAM": st.text_input("🔍 RAM"),
        "Hard Drive": st.text_input("🔍 Hard Drive"),
        "Remote Mgmt Card": st.text_input("🔍 Remote Mgmt Card"),
        "Drive Controller": st.text_input("🔍 Drive Controller"),
        "Others": st.text_input("Search term for Others")
    }

    st.markdown("---")
    st.subheader("🔍 Others Search")
    selected_component_type = st.selectbox("Select Component Category", options=[""] + component_types)
    quantity_filter = st.number_input("Minimum Quantity", min_value=0, value=0)
    status_filter = st.text_input("Hardware Status (optional)")
    fuzzy_threshold = st.slider("Fuzzy match threshold", min_value=70, max_value=100, value=85)
    selected_sheets = st.multiselect("Select sheets to search", options=xls.sheet_names, default=xls.sheet_names)

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
        results = {}
        for label, term in search_terms.items():
            if label == "Others" and not term:
                continue
            if term:
                st.header(f"🔍 Results for {label}: {term}")
                site_results = {}
                for sheet_name in selected_sheets:
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
                    except Exception as e:
                        st.warning(f"Skipping sheet '{sheet_name}' due to an error: {e}")
                        continue

                    if "general search" in df.iloc[0].astype(str).str.lower().tolist():
                        df = df[1:]

                    site_headers = df.columns.tolist()[1:]
                    for site in site_headers:
                        if site in df.columns:
                            matched_rows = df[df[site].astype(str).apply(
                                lambda x: fuzz.partial_ratio(term.lower(), str(x).lower()) >= fuzzy_threshold
                            )].copy()

                            if label == "Others":
                                if selected_component_type and "Component Type" in matched_rows.columns:
                                    matched_rows = matched_rows[matched_rows["Component Type"] == selected_component_type]
                                if quantity_filter > 0 and "Quantity" in matched_rows.columns:
                                    matched_rows = matched_rows[matched_rows["Quantity"] >= quantity_filter]
                                if status_filter and "Hardware Status" in matched_rows.columns:
                                    matched_rows = matched_rows[
                                        matched_rows["Hardware Status"].astype(str).str.contains(status_filter, case=False)
                                    ]

                            if "Link" not in matched_rows.columns:
                                matched_rows["Link"] = ""
                            matched_rows["Link"] = matched_rows[site].apply(extract_link)

                            for col in ["Component Type", "Quantity", "Hardware Status", "Location", "Notes"]:
                                if col not in matched_rows.columns:
                                    matched_rows[col] = ""

                            site_results[site] = matched_rows[[
                                site, "Component Type", "Quantity", "Hardware Status", "Location", "Notes", "Link"
                            ]]

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
                            st.markdown(f"### 📍 Site: {site}")
                            with st.expander("Expand/Minimize Results"):
                                st.dataframe(data, hide_index=True, use_container_width=True)
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
