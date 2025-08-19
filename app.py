import streamlit as st
import pandas as pd
import re
from fuzzywuzzy import fuzz
import streamlit.components.v1 as components

st.set_page_config(layout="wide")
st.title("🔍 IBM Component Multi-Search Viewer")
st.markdown("Upload the Excel file and search for multiple components to view their associated links.")

# Quick Links Section
st.markdown("---")
st.subheader("Quick Links")
col1, col2, _ = st.columns(3)

with col1:
    st.markdown("##### INVENTORY SERVERS")
    for site in ["OSA21", "OSA22", "OSA23"]:
        if st.button(site):
            components.html("", height=0)

with col2:
    st.markdown("##### FORMAT DRIVES")
    for site in ["OSA21", "OSA22", "OSA23"]:
        if st.button(site + " ", key=f"format_{site.lower()}"):
            components.html("", height=0)

st.markdown("---")

# File Upload
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
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

    # Search Inputs
    search_terms = {
        "Processor": st.text_input("🔍 Processor"),
        "RAM": st.text_input("🔍 RAM"),
        "Hard Drive": st.text_input("🔍 Hard Drive"),
        "Remote Mgmt Card": st.text_input("🔍 Remote Mgmt Card"),
        "Drive Controller": st.text
