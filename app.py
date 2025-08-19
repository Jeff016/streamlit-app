import streamlit as st
import pandas as pd
from rapidfuzz import process
from io import BytesIO

st.set_page_config(layout="wide")
st.title("🔍 IBM Component Multi-Search Viewer")
st.markdown("Upload the Excel file and search for multiple components to view their associated links.")

# File upload
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
        st.success(f"Successfully loaded: {uploaded_file.name}")
    except Exception as e:
        st.error(f"Error loading the Excel file: {e}. Please ensure it is a valid .xlsx file.")
        st.stop()

    # Search inputs
    search_term = st.text_input("🔍 Enter component name to search")
    fuzzy_threshold = st.slider("Fuzzy match threshold", min_value=70, max_value=100, value=85)
    selected_sheets = st.multiselect("Select sheets to search", options=list(xls.keys()), default=list(xls.keys()))

    if search_term and st.button("🔎 Search Components"):
        export_data = []
        for sheet_name in selected_sheets:
            df = xls[sheet_name]
            df.columns = df.columns.str.strip()
            if df.empty:
                continue

            choices = df.iloc[:, 0].astype(str).tolist()
            matches = process.extract(search_term, choices, limit=10, score_cutoff=fuzzy_threshold)

            matched_values = [match[0] for match in matches]
            filtered_df = df[df.iloc[:, 0].astype(str).isin(matched_values)]

            if not filtered_df.empty:
                st.subheader(f"📄 Sheet: {sheet_name}")
                st.dataframe(filtered_df, use_container_width=True)
                export_data.append((sheet_name, filtered_df))

        # Export results
        if export_data:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, data in export_data:
                    data.to_excel(writer, index=False, sheet_name=sheet_name)
            st.download_button(
                label="📥 Download Matching Results",
                data=output.getvalue(),
                file_name="matching_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No matching components found.")
else:
    st.info("Please upload an Excel file to begin.")
