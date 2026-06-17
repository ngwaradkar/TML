import pandas as pd
import re

f = 'DPT_PLAN-VIN_GENERATION_REPORT_06_16_2026 06_30_00.xls'

# Read raw HTML
with open(f, 'rb') as fh:
    raw = fh.read().decode('utf-8', errors='replace')

print("File length:", len(raw))
print("\nFirst 2000 chars:")
print(raw[:2000])
print("\n\nLast 500 chars:")
print(raw[-500:])

# Try all tables
tables = pd.read_html(f, header=None)
print(f"\nTables found: {len(tables)}")
for i, t in enumerate(tables):
    print(f"Table {i}: shape={t.shape}")
    # Show first non-null rows
    non_null = t.dropna(how='all')
    print(f"  Non-null rows: {len(non_null)}")
    print(non_null.head(5).to_string())
