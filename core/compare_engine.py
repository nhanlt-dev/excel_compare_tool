import pandas as pd
import numpy as np
from decimal import Decimal, InvalidOperation
import re
import unicodedata

NUMBER_RE = re.compile(r'^[\+\-]?\d+([.,]\d+)?([eE][\+\-]?\d+)?$')

def _strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join([c for c in nfkd if not unicodedata.combining(c)])

def normalize_value(v, case_sensitive: bool=False, remove_accents: bool=False):
    # None/NaN -> ""
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    s = str(v).strip()
    if s.lower() in ("NotaNumber", "none", ""):
        return ""
    # handle thousands / decimal separators
    s2 = s.replace('\xa0','')
    if ',' in s2 and '.' in s2:
        # if '.' after ',', likely decimal dot, remove commas
        if s2.rfind('.') > s2.rfind(','):
            s2 = s2.replace(',', '')
    elif ',' in s2 and '.' not in s2:
        parts = s2.split(',')
        if len(parts[-1]) <= 3 and len(parts)==2:
            s2 = s2.replace(',', '.')
        else:
            s2 = s2.replace(',', '')
    s2 = s2.replace('\u2212','-')
    s2_clean = s2.strip()
    if NUMBER_RE.match(s2_clean):
        try:
            d = Decimal(s2_clean)
            if d == d.to_integral():
                norm = format(d.quantize(1), 'f')
            else:
                norm = format(d.normalize(), 'f')
            return norm
        except (InvalidOperation, ValueError):
            pass
    text = s2_clean
    if not case_sensitive:
        text = text.lower()
    if remove_accents:
        text = _strip_accents(text)
    return text

def compare_tables(file_a, file_b, key_a, key_b, pairs, extra_a=None, extra_b=None, case_sensitive=False, remove_accents=False, preview_limit=200):
    extra_a = extra_a or []
    extra_b = extra_b or []
    pairs = pairs or []

    dfA = pd.read_excel(file_a, dtype=object, engine="openpyxl")
    dfB = pd.read_excel(file_b, dtype=object, engine="openpyxl")

    if key_a not in dfA.columns:
        raise ValueError(f"Key '{key_a}' not in file A")
    if key_b not in dfB.columns:
        raise ValueError(f"Key '{key_b}' not in file B")

    dfA['_key_norm'] = dfA[key_a].apply(lambda x: normalize_value(x, case_sensitive, remove_accents))
    dfB['_key_norm'] = dfB[key_b].apply(lambda x: normalize_value(x, case_sensitive, remove_accents))

    merged = pd.merge(dfA, dfB, on='_key_norm', how='outer', suffixes=('_A','_B'), indicator=True)

    status = []
    detail = []

    # precompute normalized series for all columns we will compare
    norm_cache = {}
    def norm_series(side, col):
        k = (side,col)
        if k in norm_cache:
            return norm_cache[k]
        if side == 'A':
            if col in merged.columns:
                s = merged[col]
            elif col + '_A' in merged.columns:
                s = merged[col + '_A']
            else:
                s = pd.Series([""]*len(merged), index=merged.index)
        else:
            if col in merged.columns:
                s = merged[col]
            elif col + '_B' in merged.columns:
                s = merged[col + '_B']
            else:
                s = pd.Series([""]*len(merged), index=merged.index)
        normed = s.map(lambda x: normalize_value(x, case_sensitive, remove_accents))
        norm_cache[k] = normed
        return normed

    for i, r in merged.iterrows():
        ind = r['_merge']
        if ind == 'left_only':
            status.append("Chỉ bên A"); detail.append("")
            continue
        if ind == 'right_only':
            status.append("Chỉ bên B"); detail.append("")
            continue
        diffs = []
        for a_col, b_col in pairs:
            na = norm_series('A', a_col).iat[i]
            nb = norm_series('B', b_col).iat[i]
            if na != nb:
                rawA = r[a_col] if a_col in merged.columns else r.get(a_col+'_A','')
                rawB = r[b_col] if b_col in merged.columns else r.get(b_col+'_B','')
                diffs.append(f"{a_col}≠{b_col}: '{rawB}'→'{rawA}'")
        if diffs:
            status.append("Khác"); detail.append("; ".join(diffs))
        else:
            status.append("Khớp"); detail.append("")
    merged['Trạng thái'] = status
    merged['Chi tiết'] = detail

    # Build output columns: keyA/keyB readable then extras then status/detail
    out_cols = []
    if key_a in merged.columns:
        out_cols.append(key_a)
    elif key_a + '_A' in merged.columns:
        out_cols.append(key_a + '_A')
    if key_b in merged.columns and key_b not in out_cols:
        out_cols.append(key_b)
    elif key_b + '_B' in merged.columns and (key_b + '_B') not in out_cols:
        out_cols.append(key_b + '_B')

    for c in extra_a:
        if c in merged.columns:
            out_cols.append(c)
        elif c + '_A' in merged.columns:
            out_cols.append(c + '_A')
    for c in extra_b:
        if c in merged.columns:
            out_cols.append(c)
        elif c + '_B' in merged.columns:
            out_cols.append(c + '_B')

    out_cols += ['Trạng thái', 'Chi tiết']
    out_cols = [c for c in out_cols if c in merged.columns]
    result_df = merged.loc[:, out_cols].copy()
    preview_df = result_df.head(preview_limit).copy()
    return result_df, preview_df
