"""
Indeed Client Configuration for 33-Column Standardized Format
Specific merge logic for Indeed's XTM, TOS, and Edit Distance files
"""

def get_indeed_merge_config():
    """Returns the Indeed-specific configuration for merging files into 33-column format"""
    
    return {
        "Indeed_Standard": {
            # Define which columns to take from each file
            "column_mapping": {
                # Columns A-I: From XTM file (Indeed's project data)
                "A_Status": {"value": "Translated", "type": "static"},
                "B_Project_ID": {"source": "XTM", "column": "Project ID", "type": "direct"},
                "C_Project_Name": {"source": "XTM", "column": "Project name", "type": "direct"},
                "D_Creation_Date": {"source": "XTM", "column": "Creation date", "type": "direct"},
                "E_Due_Date": {"source": "XTM", "column": "Due date", "type": "direct"},
                "F_Department": {"source": "XTM", "column": "Department_Indeed", "type": "direct"},
                "G_Team": {"source": "XTM", "column": "Team_Indeed", "type": "direct"},
                "H_Source_Language": {"source": "XTM", "column": "Source language", "type": "direct"},
                "I_Target_Language": {"source": "XTM", "column": "Target language", "type": "direct"},
                
                # Columns J-L: From TOS file (Indeed's order data)
                "J_Service_Type": {"source": "TOS", "column": "service_type", "type": "direct"},
                "K_Requested_By": {"source": "TOS", "column": "requested_by", "type": "direct"},
                "L_Tags": {"source": "TOS", "column": "tags", "type": "direct"},
                
                # Columns M-T: From Edit Distance file (Indeed's word counts)
                "M_No_Match": {"source": "EDIT", "column": "No match (after MTPE discount)", "type": "direct"},
                "N_50_74_Match": {"source": "EDIT", "column": "50%-74%", "type": "direct"},
                "O_75_84_Match": {"source": "EDIT", "column": "75%-84%", "type": "direct"},
                "P_85_94_Match": {"source": "EDIT", "column": "85%-94%", "type": "direct"},
                "Q_Total_Words": {"type": "formula", "formula": "sum_columns", "columns": ["M_No_Match", "N_50_74_Match", "O_75_84_Match", "P_85_94_Match"]},
                "R_95_99_Match": {"source": "EDIT", "column": "95%-99%", "type": "direct"},
                "S_100_Match": {"source": "EDIT", "column": "100%", "type": "direct"},
                "T_Repetitions": {"source": "EDIT", "column": "Repetitions", "type": "direct"},
                
                # Columns U-AG: Indeed-specific additional fields (customize as needed)
                "U_Field": {"value
