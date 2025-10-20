"""
Google Sheets Handler for Indeed
Manages Google Sheets API integration and file upload
"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import streamlit as st
from datetime import datetime
import json

class GoogleSheetsHandler:
    """Handles all Google Sheets operations for Indeed"""
    
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
    
    def create_indeed_sheet(self, dataframe, sheet_name=None):
        """
        Create a new Google Sheet with Indeed's merged data
        
        Args:
            dataframe: Pandas DataFrame with 33 columns
            sheet_name: Optional custom name for the sheet
        
        Returns:
            dict: Contains sheet URL and ID, or error information
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
            
            # Get the first worksheet
            worksheet = spreadsheet.sheet1
            
            # Rename the first worksheet
            worksheet.update_title("Indeed_Merged_Data")
            
            # Prepare data for upload
            # Convert DataFrame to list of lists (including headers)
            headers = dataframe.columns.tolist()
            values = dataframe.fillna('').values.tolist()
            all_data = [headers] + values
            
            # Calculate required sheet size
            num_rows = len(all_data)
            num_cols = len(headers)
            
            # Resize worksheet if needed
            worksheet.resize(rows=max(num_rows, 1000), cols=33)  # Indeed has 33 columns
            
            # Upload all data at once (more efficient)
            cell_range = f'A1:{self._get_column_letter(num_cols)}{num_rows}'
            worksheet.update(cell_range, all_data, value_input_option='RAW')
            
            # Format the header row
            header_format = {
                "backgroundColor": {"red": 0.2, "green": 0.4, "blue": 0.8},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "horizontalAlignment": "CENTER"
            }
            worksheet.format(f'A1:{self._get_column_letter(num_cols)}1', header_format)
            
            # Auto-resize columns for better readability
            worksheet.columns_auto_resize(0, num_cols - 1)
            
            # Set sharing permissions (anyone with link can view)
            spreadsheet.share('', perm_type='anyone', role='reader', with_link=True)
            
            return {
                "success": True,
                "url": spreadsheet.url,
                "id": spreadsheet.id,
                "name": sheet_name,
                "rows": num_rows - 1,  # Excluding header
                "columns": num_cols
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"Failed to create Google Sheet: {str(e)}"
            }
    
    def update_existing_sheet(self, sheet_url, dataframe):
        """
        Update an existing Google Sheet with new Indeed data
        
        Args:
            sheet_url: URL of the existing sheet
            dataframe: New data to upload
        
        Returns:
            dict: Success status and information
        """
        if not self.is_connected:
            if not self.connect():
                return {
                    "success": False,
                    "error": self.error_message
                }
        
        try:
            # Open existing spreadsheet
            spreadsheet = self.client.open_by_url(sheet_url)
            worksheet = spreadsheet.sheet1
            
            # Clear existing content
            worksheet.clear()
            
            # Upload new data
            headers = dataframe.columns.tolist()
            values = dataframe.fillna('').values.tolist()
            all_data = [headers] + values
            
            worksheet.update('A1', all_data, value_input_option='RAW')
            
            return {
                "success": True,
                "message": "Sheet updated successfully",
                "rows": len(values)
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": f"Failed to update sheet: {str(e)}"
            }
    
    def list_recent_sheets(self, limit=10):
        """
        List recent Indeed sheets created by this service account
        
        Args:
            limit: Maximum number of sheets to return
        
        Returns:
            list: Recent spreadsheets with Indeed in the name
        """
        if not self.is_connected:
            if not self.connect():
                return []
        
        try:
            # List all available spreadsheets
            all_files = self.client.list_spreadsheet_files()
            
            # Filter for Indeed-related sheets
            indeed_sheets = []
            for file in all_files[:limit * 2]:  # Check more files to find Indeed ones
                if 'Indeed' in file['name'] or 'indeed' in file['name']:
                    indeed_sheets.append({
                        'name': file['name'],
                        'id': file['id'],
                        'url': f"https://docs.google.com/spreadsheets/d/{file['id']}",
                        'created': file.get('createdTime', 'Unknown')
                    })
                    if len(indeed_sheets) >= limit:
                        break
            
            return indeed_sheets
            
        except Exception as e:
            st.error(f"Failed to list sheets: {str(e)}")
            return []
    
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
