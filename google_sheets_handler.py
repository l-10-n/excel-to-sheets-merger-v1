"""
Google Sheets Handler for Indeed
Enhanced with multi-tab support and comprehensive reporting
"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st
from datetime import datetime
import json

class GoogleSheetsHandler:
    """Handles all Google Sheets operations for Indeed with multi-tab support"""
    
    def __init__(self):
        """Initialize Google Sheets connection"""
        self.client = None
        self.is_connected = False
        self.error_message = None
        
    def connect(self):
        """
        Connect to Google Sheets API using service account credentials
        Returns: bool indicating success
        """
        try:
            # Define required scopes
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive.file',
                'https://www.googleapis.com/auth/drive'
            ]
            
            # Get credentials from Streamlit secrets
            if 'gcp_service_account' not in st.secrets:
                self.error_message = "Google Cloud credentials not found in secrets"
                return False
            
            # Create credentials from secrets
            creds_dict = dict(st.secrets["gcp_service_account"])
            
            # Create credentials object
            credentials = Credentials.from_service_account_info(
                creds_dict,
                scopes=scopes
            )
            
            # Authorize gspread client
            self.client = gspread.authorize(credentials)
            self.is_connected = True
            return True
            
        except Exception as e:
            self.error_message = f"Failed to connect to Google Sheets: {str(e)}"
            self.is_connected = False
            return False
    
    def create_indeed_sheet_multi_tab(self, merged_df, xtm_df, tos_df, edit_df, validation_issues, sheet_name=None):
        """
        Create a comprehensive Google Sheet with 5 tabs:
        1. Merged Report (33 columns)
        2. Validation Log
        3. XTM Raw Data
        4. TOS Raw Data
        5. Edit Distance Raw Data
        
        Args:
            merged_df: The 33-column merged DataFrame
            xtm_df: Original XTM DataFrame
            tos_df: Original TOS DataFrame
            edit_df: Original Edit Distance DataFrame
            validation_issues: List of validation warnings/issues
            sheet_name: Optional custom name for the sheet
        
        Returns:
            dict: Contains success status, URL, and other metadata
        """
        if not self.is_connected:
            if not self.connect():
                return {
                    "success": False,
                    "error": self.error_message
                }
        
        try:
            # Generate sheet name if not provided
            if not sheet_name:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                sheet_name = f"Indeed_Report_{timestamp}"
            
            # Create new spreadsheet
            spreadsheet = self.client.create(sheet_name)
            
            # Tab 1: Merged Report (rename the default sheet1)
            merged_sheet = spreadsheet.sheet1
            merged_sheet.update_title("Merged_Report")
            self._upload_dataframe_to_sheet(merged_sheet, merged_df)
            self._format_professional_header(merged_sheet, len(merged_df.columns))
            
            # Tab 2: Validation Log
            validation_sheet = spreadsheet.add_worksheet("Validation_Log", rows=500, cols=10)
            self._create_validation_log(validation_sheet, validation_issues, merged_df, xtm_df, tos_df, edit_df)
            
            # Tab 3: XTM Raw Data
            xtm_rows = max(len(xtm_df) + 1, 100)
            xtm_cols = max(len(xtm_df.columns), 20)
            xtm_sheet = spreadsheet.add_worksheet("XTM_Raw", rows=xtm_rows, cols=xtm_cols)
            self._upload_dataframe_to_sheet(xtm_sheet, xtm_df)
            self._format_raw_data_header(xtm_sheet, len(xtm_df.columns), "#e3f2fd")  # Light blue
            
            # Tab 4: TOS Raw Data
            tos_rows = max(len(tos_df) + 1, 100)
            tos_cols = max(len(tos_df.columns), 20)
            tos_sheet = spreadsheet.add_worksheet("TOS_Raw", rows=tos_rows, cols=tos_cols)
            self._upload_dataframe_to_sheet(tos_sheet, tos_df)
            self._format_raw_data_header(tos_sheet, len(tos_df.columns), "#f0fdf4")  # Light green
            
            # Tab 5: Edit Distance Raw Data
            edit_rows = max(len(edit_df) + 1, 100)
            edit_cols = max(len(edit_df.columns), 20)
            edit_sheet = spreadsheet.add_worksheet("Edit_Distance_Raw", rows=edit_rows, cols=edit_cols)
            self._upload_dataframe_to_sheet(edit_sheet, edit_df)
            self._format_raw_data_header(edit_sheet, len(edit_df.columns), "#fef3c7")  # Light yellow
            
            # Set sharing permissions (anyone with link can view)
            spreadsheet.share('', perm_type='anyone', role='reader', with_link=True)
            
            # Calculate total rows across all sheets
            total_rows = len(merged_df) + len(xtm_df) + len(tos_df) + len(edit_df)
            
            return {
                "success": True,
                "url": spreadsheet.url,
                "id": spreadsheet.id,
                "name": sheet_name,
                "tabs_created": 5,
                "total_rows": total_rows,
                "merged_rows": len(merged_df),
                "merged_columns": len(merged_df.columns)
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"Failed to create Google Sheet: {str(e)}"
            }
    
    def _upload_dataframe_to_sheet(self, worksheet, df):
        """
        Upload a DataFrame to a worksheet efficiently
        
        Args:
            worksheet: gspread worksheet object
            df: pandas DataFrame to upload
        """
        # Prepare data (convert DataFrame to list of lists with headers)
        headers = df.columns.tolist()
        values = df.fillna('').astype(str).values.tolist()
        all_data = [headers] + values
        
        # Calculate range
        num_rows = len(all_data)
        num_cols = len(headers)
        
        if num_rows > 0 and num_cols > 0:
            # Update in one batch for efficiency
            cell_range = f'A1:{self._get_column_letter(num_cols)}{num_rows}'
            worksheet.update(cell_range, all_data, value_input_option='RAW')
    
    def _create_validation_log(self, worksheet, validation_issues, merged_df, xtm_df, tos_df, edit_df):
        """
        Create a comprehensive validation log in the worksheet
        
        Args:
            worksheet: gspread worksheet for validation log
            validation_issues: List of validation warnings
            merged_df, xtm_df, tos_df, edit_df: DataFrames for statistics
        """
        log_data = []
        
        # Header
        log_data.append(["Validation Report", "", "", ""])
        log_data.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "", ""])
        log_data.append(["", "", "", ""])
        
        # File Statistics
        log_data.append(["FILE STATISTICS", "", "", ""])
        log_data.append(["File Type", "Records", "Columns", "Status"])
        log_data.append(["XTM Export", str(len(xtm_df)), str(len(xtm_df.columns)), "✓ Loaded"])
        log_data.append(["TOS Export", str(len(tos_df)), str(len(tos_df.columns)), "✓ Loaded"])
        log_data.append(["Edit Distance", str(len(edit_df)), str(len(edit_df.columns)), "✓ Loaded"])
        log_data.append(["Merged Output", str(len(merged_df)), str(len(merged_df.columns)), "✓ Created"])
        log_data.append(["", "", "", ""])
        
        # Validation Issues
        log_data.append(["VALIDATION WARNINGS", "", "", ""])
        log_data.append(["Warning Type", "Description", "Impact", "Action"])
        
        if validation_issues:
            for issue in validation_issues:
                log_data.append(["Warning", issue, "Non-critical", "Review if needed"])
        else:
            log_data.append(["None", "All validations passed", "None", "No action required"])
        
        log_data.append(["", "", "", ""])
        
        # Column Mapping Status
        log_data.append(["COLUMN MAPPING STATUS", "", "", ""])
        log_data.append(["Output Column", "Source", "Status", "Notes"])
        
        # Add column mapping details (first 10 as example)
        for col in merged_df.columns[:10]:
            if merged_df[col].notna().any():
                log_data.append([col, "Mapped", "✓", f"{merged_df[col].notna().sum()} values"])
            else:
                log_data.append([col, "Empty", "⚠", "No data mapped"])
        
        # Upload to sheet
        if log_data:
            worksheet.update('A1', log_data, value_input_option='RAW')
            
            # Format headers
            header_format = {
                "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95},
                "textFormat": {"bold": True}
            }
            worksheet.format('A1:D1', header_format)
            worksheet.format('A4:D4', header_format)
            worksheet.format('A11:D11', header_format)
            worksheet.format('A18:D18', header_format)
    
    def _format_professional_header(self, worksheet, num_cols):
        """
        Apply professional formatting to the merged report header
        
        Args:
            worksheet: gspread worksheet object
            num_cols: Number of columns to format
        """
        try:
            header_format = {
                "backgroundColor": {"red": 0.0, "green": 0.4, "blue": 0.8},  # Indeed blue
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1}
                },
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "borders": {
                    "bottom": {
                        "style": "SOLID",
                        "width": 2,
                        "color": {"red": 0, "green": 0.3, "blue": 0.6}
                    }
                }
            }
            
            # Apply formatting to header row
            worksheet.format(f'A1:{self._get_column_letter(num_cols)}1', header_format)
            
            # Freeze the header row
            worksheet.freeze(rows=1)
            
            # Auto-resize columns (if not too many)
            if num_cols <= 50:
                worksheet.columns_auto_resize(0, num_cols - 1)
                
        except Exception:
            # Formatting is optional, don't fail if it doesn't work
            pass
    
    def _format_raw_data_header(self, worksheet, num_cols, bg_color_hex):
        """
        Apply formatting to raw data sheet headers with custom color
        
        Args:
            worksheet: gspread worksheet object
            num_cols: Number of columns to format
            bg_color_hex: Hex color for background (e.g., "#e3f2fd")
        """
        try:
            # Convert hex to RGB (0-1 scale)
            r = int(bg_color_hex[1:3], 16) / 255
            g = int(bg_color_hex[3:5], 16) / 255
            b = int(bg_color_hex[5:7], 16) / 255
            
            header_format = {
                "backgroundColor": {"red": r, "green": g, "blue": b},
                "textFormat": {"bold": True},
                "horizontalAlignment": "CENTER"
            }
            
            # Apply formatting to header row
            worksheet.format(f'A1:{self._get_column_letter(num_cols)}1', header_format)
            
            # Freeze the header row
            worksheet.freeze(rows=1)
            
        except Exception:
            # Formatting is optional
            pass
    
    def _get_column_letter(self, col_num):
        """
        Convert column number to Excel-style letter (A, B, ... AA, AB, ...)
        
        Args:
            col_num: Column number (1-based)
        
        Returns:
            str: Column letter
        """
        letter = ''
        while col_num > 0:
            col_num -= 1
            letter = chr(col_num % 26 + ord('A')) + letter
            col_num //= 26
        return letter
    
    def test_connection(self):
        """
        Test if Google Sheets connection is working
        
        Returns:
            dict: Connection status and details
        """
        if self.connect():
            try:
                # Try to list files as a connection test
                files = self.client.list_spreadsheet_files(limit=1)
                return {
                    "connected": True,
                    "message": "Successfully connected to Google Sheets API",
                    "can_list_files": len(files) >= 0
                }
            except Exception as e:
                return {
                    "connected": False,
                    "message": f"Connected but cannot access sheets: {str(e)}",
                    "can_list_files": False
                }
        else:
            return {
                "connected": False,
                "message": self.error_message or "Failed to connect",
                "can_list_files": False
            }

# Convenience functions for Streamlit app
@st.cache_data(ttl=3600)
def get_sheets_handler():
    """Get or create a cached Google Sheets handler"""
    return GoogleSheetsHandler()

def display_connection_status(handler):
    """Display Google Sheets connection status in Streamlit"""
    status = handler.test_connection()
    
    if status["connected"]:
        st.success("✅ Connected to Google Sheets API")
        if status["can_list_files"]:
            st.caption("Ready to create and manage Indeed reports")
    else:
        st.error("❌ Google Sheets not connected")
        st.caption(status["message"])
        
        with st.expander("🔧 Setup Instructions"):
            st.markdown("""
            **To enable Google Sheets integration:**
            
            1. Create a Google Cloud Project
            2. Enable Google Sheets API and Google Drive API
            3. Create a Service Account and download the JSON key
            4. Add the credentials to Streamlit secrets
            
            See the README for detailed instructions.
            """)
