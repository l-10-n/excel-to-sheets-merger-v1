"""
Merge Processor for Indeed Excel Files
Enhanced with case-insensitive matching and flexible validation
"""

import pandas as pd
import numpy as np
from datetime import datetime
from indeed_config import get_indeed_merge_config, get_indeed_column_names

def normalize_column_names(df):
    """
    Normalize column names to be case-insensitive
    Convert all column names to lowercase and strip whitespace
    
    Args:
        df: pandas DataFrame
    
    Returns:
        DataFrame with normalized column names
    """
    df = df.copy()
    df.columns = df.columns.str.lower().str.strip().str.replace(' ', '_')
    return df

def find_column_fuzzy(df, target_column):
    """
    Find a column using fuzzy matching (case-insensitive)
    
    Args:
        df: pandas DataFrame
        target_column: Column name to search for
    
    Returns:
        Actual column name if found, None otherwise
    """
    target_lower = target_column.lower().strip().replace(' ', '_')
    
    for col in df.columns:
        if col.lower().strip().replace(' ', '_') == target_lower:
            return col
    
    return None

def merge_indeed_files(xtm_df, tos_df, edit_df, client_profile="Indeed_Standard"):
    """
    Merge three Excel/CSV files according to Indeed's specifications
    Now with case-insensitive column matching
    
    Args:
        xtm_df: DataFrame from XTM export
        tos_df: DataFrame from TOS export  
        edit_df: DataFrame from Edit Distance export
        client_profile: Configuration profile to use
    
    Returns:
        DataFrame with 33 columns in Indeed's format
    """
    
    # Normalize all column names for case-insensitive matching
    xtm_df_norm = normalize_column_names(xtm_df)
    tos_df_norm = normalize_column_names(tos_df)
    edit_df_norm = normalize_column_names(edit_df)
    
    # Get Indeed's configuration
    config = get_indeed_merge_config()[client_profile]
    column_mapping = config["column_mapping"]
    join_keys = config["join_keys"]
    
    # Step 1: Merge XTM and TOS data with fuzzy column matching
    merged_df = xtm_df_norm.copy()
    
    if "xtm_tos" in join_keys:
        xtm_key = join_keys["xtm_tos"]["xtm_key"].lower().replace(' ', '_')
        tos_key = join_keys["xtm_tos"]["tos_key"].lower().replace(' ', '_')
        
        # Find actual column names
        xtm_col = find_column_fuzzy(xtm_df_norm, xtm_key)
        tos_col = find_column_fuzzy(tos_df_norm, tos_key)
        
        if xtm_col and tos_col:
            # Perform left join (keep all XTM records)
            merged_df = pd.merge(
                xtm_df_norm,
                tos_df_norm,
                left_on=xtm_col,
                right_on=tos_col,
                how='left',
                suffixes=('_xtm', '_tos')
            )
    
    # Step 2: Merge with Edit Distance data
    if "xtm_edit" in join_keys:
        xtm_key = join_keys["xtm_edit"]["xtm_key"].lower().replace(' ', '_')
        edit_key = join_keys["xtm_edit"]["edit_key"].lower().replace(' ', '_')
        
        # Find actual column names
        xtm_col = find_column_fuzzy(merged_df, xtm_key)
        edit_col = find_column_fuzzy(edit_df_norm, edit_key)
        
        if xtm_col and edit_col:
            merged_df = pd.merge(
                merged_df,
                edit_df_norm,
                left_on=xtm_col,
                right_on=edit_col,
                how='left',
                suffixes=('', '_edit')
            )
    
    # Step 3: Create the 33-column output with flexible column mapping
    result_df = pd.DataFrame()
    
    for col_name, mapping in column_mapping.items():
        if mapping["type"] == "static":
            # Static value for all rows
            result_df[col_name] = mapping["value"]
            
        elif mapping["type"] == "direct":
            # Direct mapping from source file with fuzzy matching
            source = mapping["source"]
            source_col = mapping["column"].lower().replace(' ', '_')
            
            # Try to find the column in the merged dataframe
            actual_col = find_column_fuzzy(merged_df, source_col)
            
            if actual_col:
                result_df[col_name] = merged_df[actual_col]
            else:
                # Column not found, use empty values but continue
                result_df[col_name] = ""
                    
        elif mapping["type"] == "formula":
            # Calculate formulas
            if mapping["formula"] == "sum_columns":
                columns_to_sum = mapping["columns"]
                result_df[col_name] = 0
                for sum_col in columns_to_sum:
                    if sum_col in result_df.columns:
                        # Convert to numeric, treating errors as 0
                        numeric_col = pd.to_numeric(result_df[sum_col], errors='coerce').fillna(0)
                        result_df[col_name] += numeric_col
    
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
    Validate that the uploaded files have expected structure
    Now returns warnings instead of errors - processing continues
    
    Args:
        xtm_df: XTM export DataFrame
        tos_df: TOS export DataFrame
        edit_df: Edit Distance DataFrame
    
    Returns:
        list: List of validation warnings (not blocking errors)
    """
    warnings = []
    
    # Normalize column names for checking
    xtm_norm = normalize_column_names(xtm_df)
    tos_norm = normalize_column_names(tos_df)
    edit_norm = normalize_column_names(edit_df)
    
    config = get_indeed_merge_config()["Indeed_Standard"]
    validation = config.get("validation", {})
    
    # Check XTM columns (warnings only)
    required_xtm = validation.get("required_xtm_columns", [])
    for col in required_xtm:
        col_norm = col.lower().replace(' ', '_')
        if not find_column_fuzzy(xtm_norm, col_norm):
            warnings.append(f"XTM file may be missing expected column: '{col}' (will use empty values)")
    
    # Check TOS columns (warnings only)
    required_tos = validation.get("required_tos_columns", [])
    for col in required_tos:
        col_norm = col.lower().replace(' ', '_')
        if not find_column_fuzzy(tos_norm, col_norm):
            warnings.append(f"TOS file may be missing expected column: '{col}' (will use empty values)")
    
    # Check Edit Distance columns (warnings only)
    required_edit = validation.get("required_edit_columns", [])
    for col in required_edit:
        col_norm = col.lower().replace(' ', '_')
        if not find_column_fuzzy(edit_norm, col_norm):
            warnings.append(f"Edit Distance file may be missing expected column: '{col}' (will use empty values)")
    
    # Check for empty dataframes (warnings only)
    if xtm_df.empty:
        warnings.append("XTM file appears to be empty")
    if tos_df.empty:
        warnings.append("TOS file appears to be empty")
    if edit_df.empty:
        warnings.append("Edit Distance file appears to be empty")
    
    # Check row count mismatches (informational)
    if len(xtm_df) != len(tos_df):
        warnings.append(f"Row count mismatch: XTM has {len(xtm_df)} rows, TOS has {len(tos_df)} rows")
    if len(xtm_df) != len(edit_df):
        warnings.append(f"Row count mismatch: XTM has {len(xtm_df)} rows, Edit Distance has {len(edit_df)} rows")
    
    # Check for potential join key issues
    join_keys = config.get("join_keys", {})
    
    if "xtm_tos" in join_keys:
        xtm_key = join_keys["xtm_tos"].get("xtm_key", "").lower().replace(' ', '_')
        tos_key = join_keys["xtm_tos"].get("tos_key", "").lower().replace(' ', '_')
        
        xtm_col = find_column_fuzzy(xtm_norm, xtm_key)
        tos_col = find_column_fuzzy(tos_norm, tos_key)
        
        if not xtm_col:
            warnings.append(f"XTM join key '{join_keys['xtm_tos'].get('xtm_key')}' not found - rows may not align properly")
        if not tos_col:
            warnings.append(f"TOS join key '{join_keys['xtm_tos'].get('tos_key')}' not found - rows may not align properly")
    
    return warnings

def preview_merge_result(result_df, num_rows=5):
    """
    Create a preview of the merge result for Indeed
    
    Args:
        result_df: Merged dataframe
        num_rows: Number of rows to preview
    
    Returns:
        dict: Preview information
    """
    # Count non-empty cells per column
    column_stats = {}
    for col in result_df.columns:
        non_empty = result_df[col].astype(str).str.strip().ne('').sum()
        column_stats[col] = {
            'non_empty': non_empty,
            'empty': len(result_df) - non_empty,
            'percentage_filled': round((non_empty / len(result_df)) * 100, 1) if len(result_df) > 0 else 0
        }
    
    return {
        "total_rows": len(result_df),
        "total_columns": len(result_df.columns),
        "column_names": result_df.columns.tolist(),
        "preview_data": result_df.head(num_rows),
        "column_statistics": column_stats,
        "has_empty_cells": result_df.isnull().sum().sum() > 0 or (result_df == '').sum().sum() > 0,
        "client": "Indeed"
    }
