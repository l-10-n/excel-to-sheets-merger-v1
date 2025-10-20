"""
Excel to Google Sheets Merger Application - Indeed Version
Main Streamlit application for consolidating Indeed's Excel exports
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Import Indeed-specific modules
from merge_processor import merge_indeed_files, validate_indeed_files, preview_merge_result
from indeed_config import CLIENT_NAME
from google_sheets_handler import GoogleSheetsHandler, display_connection_status

# Page configuration
st.set_page_config(
    page_title="Indeed - Excel to Google Sheets Merger",
    page_icon="📊",
    layout="wide"
)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'validation_errors' not in st.session_state:
    st.session_state.validation_errors = []

# Title and description
st.title(f"📊 Excel to Google Sheets Merger - {CLIENT_NAME}")
st.markdown("### Consolidate XTM, TOS, and Edit Distance exports into standardized 33-column Google Sheet")
st.divider()

# Sidebar for information and help
with st.sidebar:
    st.header("ℹ️ About")
    st.info(f"**Client:** {CLIENT_NAME}\n**Format:** 33-column standardized output")
    
    st.divider()
    st.header("📋 Required Files")
    st.markdown("""
    **1. XTM Export:**
    - Project ID
    - Project name
    - Department_Indeed
    - Team_Indeed
    
    **2. TOS Export:**
    - order_id
    - service_type
    - requested_by
    
    **3. Edit Distance:**
    - Word count columns
    - Match percentages
    """)
    
    st.divider()
    st.header("📖 Instructions")
    st.markdown("""
    1. Upload all 3 Excel files
    2. Validate the data
    3. Review the preview
    4. Process to Google Sheets
    """)

# Main content - File upload section
st.header("📁 Step 1: Upload Excel Files")

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**XTM Export File**")
    xtm_file = st.file_uploader(
        "Project management data",
        type=['xlsx', 'xls'],
        key="xtm_upload",
        help="Upload XTM export with project details"
    )
    if xtm_file:
        st.success(f"✓ {xtm_file.name}")

with col2:
    st.markdown("**TOS Export File**")
    tos_file = st.file_uploader(
        "Order/request data",
        type=['xlsx', 'xls'],
        key="tos_upload",
        help="Upload TOS export with order information"
    )
    if tos_file:
        st.success(f"✓ {tos_file.name}")

with col3:
    st.markdown("**Edit Distance Export**")
    edit_file = st.file_uploader(
        "Word count metrics",
        type=['xlsx', 'xls'],
        key="edit_upload",
        help="Upload Edit Distance export with word counts"
    )
    if edit_file:
        st.success(f"✓ {edit_file.name}")

# Validation and Preview Section
if xtm_file and tos_file and edit_file:
    st.divider()
    st.header("📊 Step 2: Validate and Preview")
    
    if st.button("🔍 Validate Files", type="secondary", use_container_width=True):
        with st.spinner("Reading and validating files..."):
            try:
                # Read Excel files
                xtm_df = pd.read_excel(xtm_file)
                tos_df = pd.read_excel(tos_file)
                edit_df = pd.read_excel(edit_file)
                
                # Store in session state
                st.session_state.xtm_data = xtm_df
                st.session_state.tos_data = tos_df
                st.session_state.edit_data = edit_df
                
                # Validate files
                is_valid, errors = validate_indeed_files(xtm_df, tos_df, edit_df)
                st.session_state.validation_errors = errors
                
                if is_valid:
                    st.success("✅ All files validated successfully!")
                    
                    # Show file information
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("XTM Records", len(xtm_df))
                    with col2:
                        st.metric("TOS Records", len(tos_df))
                    with col3:
                        st.metric("Edit Distance Records", len(edit_df))
                    
                    # Preview tabs
                    st.subheader("📋 Data Preview")
                    tab1, tab2, tab3 = st.tabs(["XTM Data", "TOS Data", "Edit Distance Data"])
                    
                    with tab1:
                        st.write(f"Shape: {xtm_df.shape[0]} rows × {xtm_df.shape[1]} columns")
                        st.dataframe(xtm_df.head(), use_container_width=True)
                    
                    with tab2:
                        st.write(f"Shape: {tos_df.shape[0]} rows × {tos_df.shape[1]} columns")
                        st.dataframe(tos_df.head(), use_container_width=True)
                    
                    with tab3:
                        st.write(f"Shape: {edit_df.shape[0]} rows × {edit_df.shape[1]} columns")
                        st.dataframe(edit_df.head(), use_container_width=True)
                else:
                    st.error("❌ Validation failed. Please check the errors below:")
                    for error in errors:
                        st.error(f"• {error}")
                    
            except Exception as e:
                st.error(f"Error reading files: {str(e)}")
    
    # Process and merge section
    if hasattr(st.session_state, 'xtm_data') and st.session_state.validation_errors == []:
        st.divider()
        st.header("🚀 Step 3: Process and Create Google Sheet")
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.info("Click below to merge files into Indeed's 33-column format")
        
        if st.button("⚡ Process and Merge Files", type="primary", use_container_width=True):
            with st.spinner("Merging files according to Indeed configuration..."):
                try:
                    # Perform the merge
                    result_df = merge_indeed_files(
                        st.session_state.xtm_data,
                        st.session_state.tos_data,
                        st.session_state.edit_data
                    )
                    
                    st.session_state.processed_data = result_df
                    
                    # Show success and preview
                    st.success(f"✅ Successfully merged files into {len(result_df.columns)}-column format!")
                    
                    # Get preview information
                    preview_info = preview_merge_result(result_df)
                    
                    # Display metrics
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Rows", preview_info["total_rows"])
                    with col2:
                        st.metric("Total Columns", preview_info["total_columns"])
                    with col3:
                        st.metric("Client", preview_info["client"])
                    with col4:
                        st.metric("Status", "Ready for Export")
                    
                    # Show merged data preview
                    st.subheader("📊 Merged Data Preview (First 10 rows)")
                    st.dataframe(result_df.head(10), use_container_width=True)
                    
                    # Download option and Google Sheets upload
                    st.divider()
                    
                    # Google Sheets Integration
                    st.subheader("☁️ Upload to Google Sheets")
                    
                    # Initialize Google Sheets handler
                    sheets_handler = GoogleSheetsHandler()
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        sheet_name = st.text_input(
                            "Google Sheet Name (optional)",
                            placeholder=f"Indeed_Report_{datetime.now().strftime('%Y%m%d')}",
                            help="Leave blank for auto-generated name"
                        )
                    
                    with col2:
                        if st.button("📤 Create Google Sheet", type="primary", use_container_width=True):
                            with st.spinner("Creating Google Sheet..."):
                                result = sheets_handler.create_indeed_sheet(
                                    result_df,
                                    sheet_name if sheet_name else None
                                )
                                
                                if result["success"]:
                                    st.success("✅ Google Sheet created successfully!")
                                    st.balloons()
                                    
                                    # Display sheet information
                                    st.info(f"""
                                    **Sheet Created!**
                                    - 📊 Name: {result['name']}
                                    - 📝 Rows: {result['rows']}
                                    - 📋 Columns: {result['columns']}
                                    """)
                                    
                                    # Display URL with copy button
                                    st.text_input(
                                        "Google Sheets URL:",
                                        value=result['url'],
                                        key="sheet_url",
                                        help="Copy this link to share with your team"
                                    )
                                    
                                    st.markdown(f"🔗 [Open Google Sheet]({result['url']})")
                                    
                                else:
                                    st.error(f"Failed to create Google Sheet: {result['error']}")
                                    st.info("Falling back to Excel download option below")
                    
                    # Test connection button
                    with st.expander("🔧 Test Google Sheets Connection"):
                        if st.button("Test Connection"):
                            display_connection_status(sheets_handler)
                    
                    # Excel download as backup
                    st.divider()
                    st.subheader("💾 Alternative: Download as Excel")
                    
                    # Convert to Excel for download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='Indeed_Merged_Data')
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_data,
                        file_name=f"Indeed_Merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"Error during merge: {str(e)}")
                    st.exception(e)

else:
    st.info("👆 Please upload all three Excel files to begin")

# Footer
st.divider()
st.caption(f"Excel to Google Sheets Merger v1.0 | Client: {CLIENT_NAME} | 33-Column Format")
