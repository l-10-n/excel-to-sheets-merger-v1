"""
Excel to Google Sheets Merger Application
Main Streamlit application for consolidating Excel exports into standardized Google Sheets
"""

import streamlit as st
import pandas as pd
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Excel to Google Sheets Merger",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 Excel to Google Sheets Merger")
st.markdown("### Consolidate multiple Excel exports into a standardized Google Sheet format")
st.divider()

# Sidebar for client selection
with st.sidebar:
    st.header("⚙️ Configuration")
    
    selected_client = st.selectbox(
        "Select Client Profile",
        options=["Standard_33_Column", "Client_2"],
        help="Choose the client configuration for data processing rules"
    )
    
    st.divider()
    st.header("📋 Instructions")
    st.markdown("""
    1. Select your client profile
    2. Upload the 3 required Excel files
    3. Click 'Process Files'
    4. Get your Google Sheets link
    """)

# Main content - File upload section
st.header("📁 Upload Excel Files")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**XTM Export File**")
    xtm_file = st.file_uploader(
        "Choose XTM file",
        type=['xlsx', 'xls'],
        key="xtm_upload"
    )

with col2:
    st.markdown("**TOS Export File**")
    tos_file = st.file_uploader(
        "Choose TOS file",
        type=['xlsx', 'xls'],
        key="tos_upload"
    )

with col3:
    st.markdown("**Edit Distance Export**")
    edit_file = st.file_uploader(
        "Choose Edit Distance file",
        type=['xlsx', 'xls'],
        key="edit_upload"
    )

# Process button
if st.button("🚀 Process Files", type="primary", use_container_width=True):
    if all([xtm_file, tos_file, edit_file]):
        with st.spinner("Processing files..."):
            try:
                # Read Excel files
                xtm_df = pd.read_excel(xtm_file)
                tos_df = pd.read_excel(tos_file)
                edit_df = pd.read_excel(edit_file)
                
                st.success("✅ Files loaded successfully!")
                
                # Show preview
                st.subheader("Preview of loaded data:")
                
                tab1, tab2, tab3 = st.tabs(["XTM Data", "TOS Data", "Edit Distance Data"])
                
                with tab1:
                    st.write(f"Rows: {len(xtm_df)}, Columns: {len(xtm_df.columns)}")
                    st.dataframe(xtm_df.head())
                
                with tab2:
                    st.write(f"Rows: {len(tos_df)}, Columns: {len(tos_df.columns)}")
                    st.dataframe(tos_df.head())
                
                with tab3:
                    st.write(f"Rows: {len(edit_df)}, Columns: {len(edit_df.columns)}")
                    st.dataframe(edit_df.head())
                
                st.info("📝 Note: Google Sheets integration will be added in the next phase")
                
            except Exception as e:
                st.error(f"Error reading files: {str(e)}")
    else:
        st.warning("⚠️ Please upload all three files before processing")

# Footer
st.divider()
st.caption("Excel to Google Sheets Merger v1.0.0")
