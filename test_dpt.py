import pandas as pd
import numpy as np

def _is_html_xls(path):
    try:
        with open(path, 'rb') as f:
            header = f.read(16)
        return header.lstrip(b'\xef\xbb\xbf').lstrip().startswith(b'<')
    except:
        return False

def _find_header_row(df_raw, keywords, max_scan=15):
    for i in range(min(max_scan, len(df_raw))):
        vals = [str(v).upper().strip() for v in df_raw.iloc[i] if pd.notna(v) and str(v) != 'nan']
        for kw in keywords:
            if any(kw in v for v in vals):
                return i
    return None

files = {
    'TCF1': 'DPT_PLAN-VIN_GENERATION_REPORT_06_16_2026 06_30_00.xls',
    'TCF2': 'TCF2_DPT-PLAN_VIN_GENERATION_REPORT_06_16_2026 06_30_00.xls',
}

for line, fname in files.items():
    print(f"\n{'='*60}")
    print(f"FILE: {fname}")
    print(f"Is HTML: {_is_html_xls(fname)}")

    if _is_html_xls(fname):
        tables = pd.read_html(fname, header=None)
        print(f"Tables found: {len(tables)}")
        for i, t in enumerate(tables):
            print(f"  Table {i}: shape={t.shape}")
        df = tables[0]
    else:
        try:
            xf = pd.ExcelFile(fname)
            print(f"Sheets: {xf.sheet_names}")
            df = pd.read_excel(fname, sheet_name=xf.sheet_names[0], header=None)
        except Exception as e:
            print(f"ERROR reading: {e}")
            continue

    print(f"Raw shape: {df.shape}")
    print(f"\nFirst 10 rows:")
    print(df.iloc[:10].to_string())

    h = _find_header_row(df, ["MARKET", "PRODUCTFAMILY", "VC", "SALES", "SHORT"])
    print(f"\nHeader row detected at: {h}")

    if h is not None:
        data = df.iloc[h+1:].copy().reset_index(drop=True)
        while len(data.columns) < 10:
            data[len(data.columns)] = np.nan
        data = data.iloc[:, :10]
        data.columns = ["MARKET","PRODUCT_FAMILY","SHORT_VC","SALES_DESC","DPT_PLAN","DPT_VIN","C6","C7","C8","C9"]
        for col in data.columns:
            if data[col].dtype == 'object':
                data[col] = data[col].astype(str).str.strip()
                data[col] = data[col].replace({"nan": np.nan, "None": np.nan, "": np.nan})
        data["DPT_PLAN"] = pd.to_numeric(data["DPT_PLAN"], errors='coerce').fillna(0).astype(int)
        data["DPT_VIN"] = pd.to_numeric(data["DPT_VIN"], errors='coerce').fillna(0).astype(int)
        data = data[data["SHORT_VC"].notna()].reset_index(drop=True)
        print(f"\nParsed rows: {len(data)}")
        print(f"DPT_VIN sum: {data['DPT_VIN'].sum()}")
        print(f"DPT_PLAN sum: {data['DPT_PLAN'].sum()}")
        print(f"\nSample rows (SHORT_VC, SALES_DESC, DPT_PLAN, DPT_VIN):")
        print(data[["SHORT_VC","SALES_DESC","DPT_PLAN","DPT_VIN"]].head(10).to_string())
