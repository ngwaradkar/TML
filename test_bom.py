import pandas as pd, numpy as np, json, os
from pathlib import Path
from datetime import datetime

DATA_DIR = Path('data')
BOM_FILE = DATA_DIR / 'bom_master.json'
BOM_SOURCE_FILE = Path('d:/TML PPC Dashboard/Bom details.xlsx')
DATA_DIR.mkdir(exist_ok=True)

df = pd.read_excel(str(BOM_SOURCE_FILE))
df.columns = [str(c).strip() for c in df.columns]
rename = {}
for c in df.columns:
    cu = c.upper().strip()
    if 'SHORT' in cu and 'VEHICLE' in cu:
        rename[c] = 'SHORT_VC'
    elif 'FRONT' in cu and 'WIRING' in cu:
        rename[c] = 'FRONT_WIRING'
    elif 'COCKPIT' in cu:
        rename[c] = 'COCKPIT'
    elif 'ENGINE' in cu:
        rename[c] = 'ENGINE'
df.rename(columns=rename, inplace=True)
needed = ['SHORT_VC', 'FRONT_WIRING', 'COCKPIT', 'ENGINE']
for n in needed:
    if n not in df.columns:
        df[n] = np.nan
df = df[needed].copy()
for col in df.columns:
    df[col] = df[col].astype(str).str.strip()
    df[col] = df[col].replace({'nan': np.nan, 'None': np.nan, '': np.nan})
df = df.dropna(subset=['SHORT_VC'])

print(f'BOM rows: {len(df)}')
print(f'Unique VCs: {df["SHORT_VC"].nunique()}')
print(df.head(3).to_string())

records = df.fillna('').to_dict('records')
with open(BOM_FILE, 'w') as f:
    json.dump({
        'source': str(BOM_SOURCE_FILE),
        'loaded_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'row_count': len(df),
        'data': records
    }, f, indent=2)
print('BOM saved to data/bom_master.json OK')
