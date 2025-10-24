import pandas as pd

def load_headers(path, engine="openpyxl"):
    """Read headers of first sheet."""
    df = pd.read_excel(path, nrows=0, engine=engine)
    return [str(c) for c in df.columns]

def read_table(path, dtype=str, engine="openpyxl"):
    """Read full sheet into DataFrame as strings to avoid type mismatch."""
    # dtype=str will coerce values to str (keeps NaN as NaN)
    return pd.read_excel(path, dtype=dtype, engine=engine)
