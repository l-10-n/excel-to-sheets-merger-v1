"""
Merge Processor for Indeed Excel Files
Handles the actual merging of XTM, TOS, and Edit Distance files
"""

import pandas as pd
import numpy as np
from datetime import datetime
from indeed_config import get_indeed_merge_config, get_indeed_column_names

def merge_indeed_files(xtm_df, tos_df, edit_df, client_profile="Indeed_Standard"):
    """
    Merge three Excel files according to Indeed's specifications
    
    Args:
        xtm_df: DataFrame from XTM export
        tos_df: DataFrame from TOS export  
        edit_df: DataFrame from Edit Distance export
        client_profile: Configuration profile to use
    
    Returns:
        DataFrame with 33 columns in Indeed's format
    """
    
    # Get Indeed's configuration
    config = get_indeed_merge_config()[client_profile]
    column_mapping = config["column_mapping"]
    join_keys = config["join_keys"]
    
    # Step 1: Merge XTM and TOS data
    if "xtm_tos" in join_keys:
        xtm_key = join_keys["xtm_tos"]["xtm_key"]
        tos_key = join_keys["xtm_tos"]["tos_key"]
        
        # Perform left join (keep all XTM records)
        merged_df = pd.merge(
            xtm_df,
            tos_df,
            left_on=xtm_key,
            right_on=tos_key,
            how='left',
            suffixes=('_xtm', '_tos')
        )
    else:
        merged_df = xtm_df.copy()
    
    # Step 2: Merge with Edit Distance data
    if "xtm_edit" in join_keys:
        xtm_key = join_keys["xtm_edit"]["xtm_key"]
        edit_key = join_keys["xtm_edit"]["edit_key"]
        
        # Check if edit_key exists in edit_df
        if edit_key in edit_df.columns:
            merged_df = pd.merge(
                merged_df,
                edit_df,
                left_on=xtm_key,
                right_on=edit_key,
                how='left',
                suffixes=('', '_edit')
            )
    
    # Step 3: Create the 33-column output
    result_df = pd.DataFrame()
    
    for col_name, mapping in column_mapping.items():
        if mapping["type"] == "static":
            # Static value for all rows
            result_df[col_name] = mapping["value"]
            
        elif mapping["type"] == "direct":
            # Direct mapping from source file
            source = mapping["source"]
            source_col = mapping["column"]
            
            if source == "XTM":
                if source_col in xtm_df.columns:
                    result_df[col_name] = merged_df[source_col] if source_col in merged_df.columns else ""
                else:
                    result_df[col_name] = ""
                    
            elif source == "TOS":
                if source_col in tos_df.columns:
                    result_df[col_name] = merged_df[source_col] if source_col in merged_df.columns else ""
                else:
                    result_df[col_name] = ""
                    
            elif source == "EDIT":
                if source_col in edit_df.columns:
                    result_df[col_name] = merged_df[source_col] if source_col in merged_df.columns else ""
                else:
                    result_df[col_name] = ""
                    
        elif mapping["type"] == "formula":
            # Calculate formulas
            if mapping["formula"] == "sum_columns":
                columns_to_sum = mapping["columns"]
                result_df[col_name] = 0
                for sum_col in columns_to_sum:
                    if sum_col in result_df.columns:
                        result_df[col_name] += pd.to_numeric(result_df[sum_col], errors='coerce').fillna(0)
    
    # Step 4: Ensure we have exactly 33 columns in the right order
    final_columns = get_indeed_column_names()
    for col in final_columns:
        if col not in result_df.columns:
            result_df[col] = ""
    
    # Reorder columns to match Indeed's specification
    result_df = result_df[final_columns]
    
    # Step 5: Clean up the data
    result_df = result_df.fillna("")
    
    return result_df

def validate_indeed_files(xtm_df, tos_df, edit_df):
    """
    Validate that the uploaded files have the required columns for Indeed
    
    Returns:
        tuple: (is_valid, error_messages)
    """
    errors = []
    config = get_indeed_merge_config()["Indeed_Standard"]
    validation = config.get("validation", {})
    
    # Check XTM columns
    required_xtm = validation.get("required_xtm_columns", [])
    missing_xtm = [col for col in required_xtm if col not in xtm_df.columns]
    if missing_xtm:
        errors.append(f"XTM file missing columns: {', '.join(missing_xtm)}")
    
    # Check TOS columns
    required_tos = validation.get("required_tos_columns", [])
    missing_tos = [col for col in required_tos if col not in tos_df.columns]
    if missing_tos:
        errors.append(f"TOS file missing columns: {', '.join(missing_tos)}")
    
    # Check Edit Distance columns
    required_edit = validation.get("required_edit_columns", [])
    missing_edit = [col for col in required_edit if col not in edit_df.columns]
    if missing_edit:
        errors.append(f"Edit Distance file missing columns: {', '.join(missing_edit)}")
    
    is_valid = len(errors) == 0
    return is_valid, errors

def preview_merge_result(result_df, num_rows=5):
    """
    Create a preview of the merge result for Indeed
    
    Args:
        result_df: Merged dataframe
        num_rows: Number of rows to preview
    
    Returns:
        dict: Preview information
    """
    return {
        "total_rows": len(result_df),
        "total_columns": len(result_df.columns),
        "column_names": result_df.columns.tolist(),
        "preview_data": result_df.head(num_rows),
        "has_empty_cells": result_df.isnull().sum().sum() > 0,
        "client": "Indeed"
    }
