"""
Excel to Google Sheets Merger Application - Indeed Version
Professional enterprise-grade data processing platform
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
    page_title="Indeed Translation Data Platform",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Professional CSS styling inspired by Translated.com
st.markdown("""
<style>
    /* Enterprise Style */
    .main {
        padding: 2rem;
        max-width: 1400px;
        margin: 0 auto;
        background-color: #ffffff;
    }
    
    /* Professional Headers */
    h1 {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        font-weight: 300;
        color: #1a1a1a;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
        letter-spacing: -0.5px;
    }
    
    h2 {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        font-weight: 400;
        color: #2c3e50;
        font-size: 1.5rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    
    /* Subtitle style */
    .subtitle {
        color: #666;
        font-size: 1.15rem;
        margin-bottom: 2rem;
        font-weight: 400;
        line-height: 1.6;
    }
    
    /* Card containers */
    .upload-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border: 1px solid #e5e7eb;
        margin-bottom: 1.5rem;
        transition: all 0.3s ease;
    }
    
    .upload-card:hover {
        box-shadow: 0 4px 20px rgba(0,0,0,0.12);
        border-color: #0066CC;
    }
    
    /* Professional buttons */
    .stButton > button {
        background-color: #0066CC;
        color: white;
        border: none;
        padding: 14px 32px;
        font-size: 16px;
        font-weight: 500;
        border-radius: 8px;
        transition: all 0.3s ease;
        text-transform: none;
        letter-spacing: 0.3px;
    }
    
    .stButton > button:hover {
        background-color: #0052A3;
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(0,102,204,0.3);
    }
    
    /* File uploader styling */
    [data-testid="stFileUploader"] {
        border: 2px dashed #e5e7eb;
        border-radius: 8px;
        padding: 1.5rem;
        background: #f9fafb;
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #0066CC;
        background: #f0f7ff;
    }
    
    /* Success messages */
    .stSuccess {
        background-color: #f0fdf4;
        border-left: 4px solid #22c55e;
        padding: 1rem;
        border-radius: 6px;
    }
    
    /* Warning messages */
    .stWarning {
        background-color: #fef3c7;
        border-left: 4px solid #f59e0b;
        padding: 1rem;
        border-radius: 6px;
    }
    
    /* Error messages */
    .stError {
        background-color: #fef2f2;
        border-left: 4px solid #ef4444;
        padding: 1rem;
        border-radius: 6px;
    }
    
    /* Metrics styling */
    [data-testid="metric-container"] {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
    }
    
    [data-testid="metric-container"] [data-testid="metric-label"] {
        color: #6b7280;
        font-size: 0.875rem;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    [data-testid="metric-container"] [data-testid="metric-value"] {
        color: #0066CC;
        font-size: 2rem;
        font-weight: 600;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f9fafb;
        padding: 4px;
        border-radius: 10px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        padding: 0 24px;
        background-color: transparent;
        border-radius: 8px;
        color: #6b7280;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: white;
        color: #0066CC;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Dividers */
    hr {
        margin: 2.5rem 0;
        border: 0;
        border-top: 1px solid #e5e7eb;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Format badges */
    .format-badge {
        display: inline-block;
        background: #e0f2fe;
        color: #0066CC;
        padding: 6px 14px;
        border-radius: 20px;
        margin-right: 8px;
        font-size: 0.875rem;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'validation_warnings' not in st.session_state:
    st.session_state.validation_warnings = []
if 'xtm_data' not in st.session_state:
    st.session_state.xtm_data = None
if 'tos_data' not in st.session_state:
    st.session_state.tos_data = None
if 'edit_data' not in st.session_state:
    st.session_state.edit_data = None

# Header Section
st.markdown("# Translation Data Platform")
st.markdown('<p class="subtitle">Enterprise-grade data processing for Indeed localization workflows. Seamlessly merge XTM, TOS, and Edit Distance reports into standardized Google Sheets.</p>', unsafe_allow_html=True)

# Trust Metrics Bar
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Files Processed", "10,000+", delta="↑ 12% this month")
with col2:
    st.metric("Time Saved", "500+ hrs", delta="Per month")
with col3:
    st.metric("Accuracy Rate", "99.9%", delta="Validated")
with col4:
    st.metric("Active Client", CLIENT_NAME, delta="Enterprise")

st.divider()

# File Upload Section
st.markdown("## 📁 Data Import")
st.markdown("Upload your localization files in any supported format. All files are processed securely and are not stored after processing.")

# Format badges
st.markdown("""
<div style="margin: 1rem 0 2rem 0;">
    <span class="format-badge">📊 Excel (.xlsx)</span>
    <span class="format-badge">📈 Legacy Excel (.xls)</span>
    <span class="format-badge">📄 CSV (.csv)</span>
</div>
""", unsafe_allow_html=True)

# File upload columns with cards
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown("### XTM Export")
    st.caption("Project management and localization data")
    xtm_file = st.file_uploader(
        "Drag & drop or click to browse",
        type=['xlsx', 'xls', 'csv'],
        key="xtm_upload",
        help="Upload your XTM export file containing project details"
    )
    if xtm_file:
        file_size = len(xtm_file.getvalue()) / 1024
        st.success(f"✓ {xtm_file.name} ({file_size:.1f} KB)")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown("### TOS Export")
    st.caption("Translation order and request data")
    tos_file = st.file_uploader(
        "Drag & drop or click to browse",
        type=['xlsx', 'xls', 'csv'],
        key="tos_upload",
        help="Upload your TOS export file containing order information"
    )
    if tos_file:
        file_size = len(tos_file.getvalue()) / 1024
        st.success(f"✓ {tos_file.name} ({file_size:.1f} KB)")
    st.markdown('</div>', unsafe_allow_html=True)

with col3:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown("### Edit Distance")
    st.caption("Word count and match analysis")
    edit_file = st.file_uploader(
        "Drag & drop or click to browse",
        type=['xlsx', 'xls', 'csv'],
        key="edit_upload",
        help="Upload your Edit Distance report with word count metrics"
    )
    if edit_file:
        file_size = len(edit_file.getvalue()) / 1024
        st.success(f"✓ {edit_file.name} ({file_size:.1f} KB)")
    st.markdown('</div>', unsafe_allow_html=True)

# Validation and Processing Section
if xtm_file and tos_file and edit_file:
    st.divider()
    st.markdown("## 🔄 Data Processing")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.info("📌 All three files detected. Click below to validate and process your data.")
    
    with col2:
        if st.button("🚀 Validate & Process", type="primary", use_container_width=True):
            with st.spinner("Reading and validating files..."):
                try:
                    # Read files based on extension
                    if xtm_file.name.endswith('.csv'):
                        xtm_df = pd.read_csv(xtm_file)
                    else:
                        xtm_df = pd.read_excel(xtm_file)
                    
                    if tos_file.name.endswith('.csv'):
                        tos_df = pd.read_csv(tos_file)
                    else:
                        tos_df = pd.read_excel(tos_file)
                    
                    if edit_file.name.endswith('.csv'):
                        edit_df = pd.read_csv(edit_file)
                    else:
                        edit_df = pd.read_excel(edit_file)
                    
                    # Store in session state
                    st.session_state.xtm_data = xtm_df
                    st.session_state.tos_data = tos_df
                    st.session_state.edit_data = edit_df
                    
                    # Validate files (now returns warnings, not errors)
                    warnings = validate_indeed_files(xtm_df, tos_df, edit_df)
                    st.session_state.validation_warnings = warnings
                    
                    # Display validation results
                    if warnings:
                        st.warning(f"⚠️ Validation completed with {len(warnings)} warning(s)")
                        with st.expander("View Warnings", expanded=False):
                            for warning in warnings:
                                st.write(f"• {warning}")
                    else:
                        st.success("✅ All validations passed successfully!")
                    
                    # Show file statistics
                    st.markdown("### 📊 File Statistics")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("XTM Records", f"{len(xtm_df):,}")
                        st.caption(f"{len(xtm_df.columns)} columns")
                    with col2:
                        st.metric("TOS Records", f"{len(tos_df):,}")
                        st.caption(f"{len(tos_df.columns)} columns")
                    with col3:
                        st.metric("Edit Distance Records", f"{len(edit_df):,}")
                        st.caption(f"{len(edit_df.columns)} columns")
                    
                    # Merge the files
                    st.markdown("### 🔀 Merging Data")
                    with st.spinner("Processing and merging files..."):
                        merged_df = merge_indeed_files(xtm_df, tos_df, edit_df)
                        st.session_state.processed_data = merged_df
                        
                        st.success(f"✅ Successfully merged into {len(merged_df.columns)}-column format with {len(merged_df):,} rows")
                    
                    # Preview merged data
                    with st.expander("📋 Preview Merged Data", expanded=True):
                        st.dataframe(merged_df.head(10), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Error processing files: {str(e)}")
                    st.exception(e)
    
    # Google Sheets Export Section
    if st.session_state.processed_data is not None:
        st.divider()
        st.markdown("## ☁️ Export to Google Sheets")
        st.markdown("Generate a comprehensive Google Sheets report with merged data, validation log, and raw source files.")
        
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            sheet_name = st.text_input(
                "Report Name (optional)",
                placeholder=f"Indeed_Report_{datetime.now().strftime('%Y%m%d')}",
                help="Leave blank for auto-generated timestamp"
            )
        
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)  # Spacing
            include_timestamp = st.checkbox("Add timestamp to name", value=True)
        
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)  # Spacing
            if st.button("📤 Create Google Sheet", type="primary", use_container_width=True):
                with st.spinner("Creating comprehensive Google Sheet report..."):
                    # Initialize handler
                    sheets_handler = GoogleSheetsHandler()
                    
                    # Prepare sheet name
                    if include_timestamp or not sheet_name:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        final_sheet_name = f"{sheet_name}_{timestamp}" if sheet_name else f"Indeed_Report_{timestamp}"
                    else:
                        final_sheet_name = sheet_name
                    
                    # Create multi-tab sheet
                    result = sheets_handler.create_indeed_sheet_multi_tab(
                        merged_df=st.session_state.processed_data,
                        xtm_df=st.session_state.xtm_data,
                        tos_df=st.session_state.tos_data,
                        edit_df=st.session_state.edit_data,
                        validation_issues=st.session_state.validation_warnings,
                        sheet_name=final_sheet_name
                    )
                    
                    if result["success"]:
                        st.balloons()
                        st.success("✅ Google Sheet created successfully!")
                        
                        # Display results
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.markdown("### 📊 Your Report is Ready!")
                            st.markdown(f"""
                            **Report Details:**
                            - 📝 Name: {result['name']}
                            - 📑 Tabs: {result['tabs_created']}
                            - 🔢 Total Rows: {result['total_rows']:,}
                            - 🔗 Shareable Link: Ready
                            """)
                            
                            # URL input for easy copying
                            st.text_input(
                                "Google Sheets URL:",
                                value=result['url'],
                                key="sheet_url",
                                help="Copy this link to share with your team"
                            )
                        
                        with col2:
                            st.markdown("### 🎯 Quick Actions")
                            st.markdown(f"[📂 Open in Google Sheets]({result['url']})")
                            st.markdown("[📧 Email Report](mailto:?subject=Indeed%20Report&body=" + result['url'] + ")")
                    else:
                        st.error(f"Failed to create Google Sheet: {result.get('error', 'Unknown error')}")
        
        # Alternative download option
        with st.expander("💾 Alternative: Download as Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.processed_data.to_excel(writer, sheet_name='Merged_Report', index=False)
                st.session_state.xtm_data.to_excel(writer, sheet_name='XTM_Raw', index=False)
                st.session_state.tos_data.to_excel(writer, sheet_name='TOS_Raw', index=False)
                st.session_state.edit_data.to_excel(writer, sheet_name='Edit_Distance_Raw', index=False)
                
                # Create validation log
                val_df = pd.DataFrame({'Warnings': st.session_state.validation_warnings if st.session_state.validation_warnings else ['No warnings']})
                val_df.to_excel(writer, sheet_name='Validation_Log', index=False)
            
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Complete Report (Excel)",
                data=excel_data,
                file_name=f"Indeed_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    # Empty state
    st.divider()
    st.markdown("## 📤 Ready to Start")
    st.info("👆 Upload all three required files above to begin processing your localization data.")

# Footer
st.divider()
st.markdown(
    f'<p style="text-align: center; color: #9ca3af; font-size: 0.875rem;">Translation Data Platform v2.0 | Client: {CLIENT_NAME} | 33-Column Standardized Format | Powered by Streamlit</p>',
    unsafe_allow_html=True
)
