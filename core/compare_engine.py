import pandas as pd
import numpy as np

def _normalize_series(s, case_sensitive: bool):
    # s is pandas Series (object)
    s = s.astype(object).fillna("")
    s = s.map(lambda x: str(x).strip())
    if not case_sensitive:
        s = s.str.lower()
    return s

def compare_tables(file_a, file_b,
                   key_a, key_b,
                   pairs,            # list of tuples (colA, colB)
                   extra_a=None, extra_b=None,
                   case_sensitive=False,
                   preview_limit=200,
                   batch_size=None):
    """
    Returns: result_df (DataFrame), preview_df (limited rows for UI)
    result_df columns: chosen key(s) and extra columns plus Trạng thái, Chi tiết
    """

    extra_a = extra_a or []
    extra_b = extra_b or []
    pairs = pairs or []

    # read full tables
    dfA = pd.read_excel(file_a, dtype=str, engine="openpyxl")
    dfB = pd.read_excel(file_b, dtype=str, engine="openpyxl")

    # Normalize keys
    if key_a not in dfA.columns:
        raise ValueError(f"Cột khóa '{key_a}' không tồn tại trong File A")
    if key_b not in dfB.columns:
        raise ValueError(f"Cột khóa '{key_b}' không tồn tại trong File B")

    dfA["_key"] = _normalize_series(dfA[key_a], case_sensitive)
    dfB["_key"] = _normalize_series(dfB[key_b], case_sensitive)

    # Merge outer to capture all keys
    merged = pd.merge(dfA, dfB, on="_key", how="outer", suffixes=("_A","_B"), indicator=True)

    # Compute status:
    # both => check pairs for differences
    statuses = []
    details = []
    # If no pairs provided, status is based on presence
    for idx, row in merged.iterrows():
        ind = row["_merge"]
        if ind == "left_only":
            statuses.append("Chỉ bên A")
            details.append("")
        elif ind == "right_only":
            statuses.append("Chỉ bên B")
            details.append("")
        else:  # both
            diffs = []
            for a_col, b_col in pairs:
                # resolve names after merge
                colA = a_col if a_col in merged.columns else (a_col + "_A" if a_col + "_A" in merged.columns else None)
                colB = b_col if b_col in merged.columns else (b_col + "_B" if b_col + "_B" in merged.columns else None)
                valA = "" if colA is None else ("" if pd.isna(row[colA]) else str(row[colA]).strip())
                valB = "" if colB is None else ("" if pd.isna(row[colB]) else str(row[colB]).strip())
                compA = valA if case_sensitive else valA.lower()
                compB = valB if case_sensitive else valB.lower()
                if compA != compB:
                    diffs.append(f"{a_col}≠{b_col}: '{valB}'→'{valA}'")
            if diffs:
                statuses.append("Khác")
                details.append("; ".join(diffs))
            else:
                statuses.append("Khớp")
                details.append("")
    merged["Trạng thái"] = statuses
    merged["Chi tiết"] = details

    # Build output column list: include readable keys and extras
    out_cols = []
    # try original key columns (prefer A then B)
    if key_a in merged.columns:
        out_cols.append(key_a)
    elif key_a + "_A" in merged.columns:
        out_cols.append(key_a + "_A")
    if key_b in merged.columns and (key_b not in out_cols):
        out_cols.append(key_b)
    elif key_b + "_B" in merged.columns and (key_b + "_B" not in out_cols):
        out_cols.append(key_b + "_B")

    # extras
    for c in extra_a:
        if c in merged.columns:
            out_cols.append(c)
        elif c + "_A" in merged.columns:
            out_cols.append(c + "_A")
    for c in extra_b:
        if c in merged.columns:
            out_cols.append(c)
        elif c + "_B" in merged.columns:
            out_cols.append(c + "_B")

    # Finally status fields
    out_cols += ["Trạng thái", "Chi tiết"]

    # Filter existing
    out_cols = [c for c in out_cols if c in merged.columns]

    result_df = merged.loc[:, out_cols].copy()

    # preview
    preview_df = result_df.head(preview_limit).copy()

    return result_df, preview_df
