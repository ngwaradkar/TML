import sys
sys.path.insert(0, '.')
import pandas as pd, numpy as np, json, os
from pathlib import Path
from datetime import datetime

# Simulate the app functions inline
DATA_DIR = Path('data')
BOM_FILE = DATA_DIR / 'bom_master.json'
BOM_SOURCE_FILE = Path('d:/TML PPC Dashboard/Bom details.xlsx')

MODEL_MAP = {
    "5497": ("PUNCH", "TCF1"),
    "5468": ("PUNCH.EV / NOVA", "TCF1"),
    "5473": ("HARRIER.EV", "TCF2"),
    "5466": ("SAFARI", "TCF2"),
    "5479": ("SAFARI", "TCF2"),
    "5605": ("SAFARI", "TCF2"),
    "5465": ("HARRIER", "TCF2"),
    "5478": ("HARRIER", "TCF2"),
    "5604": ("HARRIER", "TCF2"),
    "5483": ("SAFARI.EV", "TCF2"),
}

FLOAT_COLS = [
    "TOTAL_FLOAT","PBS_FLOAT","PBS_TO_POLISHING","POLISHING_TO_TOPCOAT",
    "TOPCOAT_TO_WETSANDING_ROOFBLACK","TOPCOAT_TO_WETSANDING_FRESH",
    "WETSANDING_TO_SEALANT","TOTAL_UPTO_SEALANT","PT_ENTRY_TO_SEALENT",
    "BIW_LIFTING_TO_PT","PT_BYPASS",
]
PF_HEADERS = [
    "MARKET","PRODUCT_FAMILY","SALES_DESCRIPTION","SHORT_VC","PACK",
    "TOTAL_FLOAT","PBS_FLOAT","PBS_TO_POLISHING","POLISHING_TO_TOPCOAT",
    "TOPCOAT_TO_WETSANDING_ROOFBLACK","TOPCOAT_TO_WETSANDING_FRESH",
    "WETSANDING_TO_SEALANT","TOTAL_UPTO_SEALANT","PT_ENTRY_TO_SEALENT",
    "BIW_LIFTING_TO_PT","PT_BYPASS","WIRING_PART_NUMBER","COCKPIT","ENGINE","FOR_MODEL_FLOAT",
]

def _is_html_xls(path):
    try:
        with open(path, 'rb') as f:
            header = f.read(16)
        return header.lstrip(b'\xef\xbb\xbf').lstrip().startswith(b'<')
    except:
        return False

def _find_header_row(df_raw, keywords, max_scan=15):
    for i in range(min(max_scan, len(df_raw))):
        vals = [str(v).upper().strip() for v in df_raw.iloc[i] if pd.notna(v)]
        for kw in keywords:
            if any(kw in v for v in vals):
                return i
    return None

f = 'PPC_Float_Report_Paint_16_06_2026 09_14_33_PM.xls'
print(f"File: {f}")
print(f"Is HTML XLS: {_is_html_xls(f)}")

# Read it
tables = pd.read_html(f, header=None)
df = tables[0]
print(f"Raw shape: {df.shape}")

h = _find_header_row(df, ["MARKET", "PRODUCT FAMILY", "SHORT VC"])
print(f"Header row index: {h}")

data = df.iloc[h + 1:].copy().reset_index(drop=True)
while len(data.columns) < 16:
    data[len(data.columns)] = np.nan
data = data.iloc[:, :16]
data.columns = PF_HEADERS[:16]

for c in ["WIRING_PART_NUMBER","COCKPIT","ENGINE","FOR_MODEL_FLOAT"]:
    data[c] = ""

for col in data.columns:
    if data[col].dtype == 'object':
        data[col] = data[col].astype(str).str.strip()
        data[col] = data[col].replace({"nan": np.nan, "None": np.nan, "": np.nan, "0": np.nan, "0.0": np.nan})

data = data[data["SHORT_VC"].notna()]
data = data[~data["SHORT_VC"].astype(str).str.contains("Total|Grand|TCF", na=False, case=False)]

for col in FLOAT_COLS:
    data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).astype(int)

def classify(vc):
    vc_str = str(vc).strip()[:4]
    for k, (mod, ln) in MODEL_MAP.items():
        if vc_str.startswith(k[:4]):
            return pd.Series([mod, ln])
    return pd.Series(["OTHER", "UNKNOWN"])

data[["MODEL","LINE"]] = data["SHORT_VC"].apply(classify)
data = data[data["LINE"] != "UNKNOWN"].reset_index(drop=True)

print(f"\nParsed rows: {len(data)}")
print(f"Lines: {data['LINE'].value_counts().to_dict()}")
print(f"Models: {data['MODEL'].value_counts().to_dict()}")
print("\nSample:")
print(data[["SHORT_VC","SALES_DESCRIPTION","MODEL","LINE","TOTAL_FLOAT","PBS_FLOAT"]].head(5).to_string())
