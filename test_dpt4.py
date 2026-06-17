import pandas as pd
from bs4 import BeautifulSoup

f = 'DPT_PLAN-VIN_GENERATION_REPORT_06_16_2026 06_30_00.xls'

with open(f, 'rb') as fh:
    raw = fh.read().decode('utf-8', errors='replace')

soup = BeautifulSoup(raw, 'html.parser')
table = soup.find('table')
rows = table.find_all('tr')
print(f"Rows in table: {len(rows)}")
print("\nFirst 5 rows content:")
for i, r in enumerate(rows[:5]):
    cells = [td.get_text(strip=True) for td in r.find_all(['td','th'])]
    print(f"  Row {i}: {cells}")

print("\nAll rows with data:")
all_rows = []
for i, r in enumerate(rows):
    cells = [td.get_text(strip=True) for td in r.find_all(['td','th'])]
    if any(c for c in cells):
        print(f"  Row {i}: {cells[:8]}...")
        all_rows.append(cells)

# Build DataFrame
import numpy as np
if len(all_rows) > 1:
    df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
    print(f"\nDataFrame shape: {df.shape}")
    print(f"Columns: {list(df.columns[:6])}")
    print(f"\nVIN column values: {df['TCF/-VIN'].tolist()}")
    print(f"\nSample:")
    print(df[['MARKET','ProductFamily','VC','SALES DESC','TCF/-Plan','TCF/-VIN']].head(10).to_string())
