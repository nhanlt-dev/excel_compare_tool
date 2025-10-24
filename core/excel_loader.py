import pandas as pd

def load_excel_columns(path):
    # read only headers
    try:
        df = pd.read_excel(path, nrows=0, engine="openpyxl")
        return [str(c) for c in df.columns]
    except Exception as e:
        raise e

def read_table(path):
    # read all as object to preserve values
    return pd.read_excel(path, dtype=object, engine="openpyxl")
