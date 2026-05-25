"""
utils.py — AIDA Ventures Integrated Dashboard
Helper functions for data cleaning and formatting.
"""
import math


def clean_val(v):
    """Convert NaN or None to empty string; leave others as-is."""
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    return v


def safe_str(v):
    """Convert value to string, returning '' for NaN/None."""
    cv = clean_val(v)
    if cv == "":
        return ""
    return str(cv).strip()


def row_to_dict(df, row_idx, key_col, val_cols):
    """Extract a row from a dataframe into a labelled dict."""
    return {col: safe_str(df.iloc[row_idx][col]) for col in val_cols}


def df_to_records(df, skip_empty=True):
    """Convert DataFrame to list of clean dicts, skipping fully-empty rows."""
    records = []
    for _, row in df.iterrows():
        d = {k: safe_str(v) for k, v in row.items()}
        if skip_empty and all(v == "" for v in d.values()):
            continue
        records.append(d)
    return records
