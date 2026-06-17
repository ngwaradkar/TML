import pandas as pd
import numpy as np

f = 'DPT_PLAN-VIN_GENERATION_REPORT_06_16_2026 06_30_00.xls'
tables = pd.read_html(f, header=None)
df = tables[0]

print("Full file contents (all rows):")
print(df.to_string())
print("\n\nAll non-null values in column 2 (VC):")
print(df.iloc[:, 2].dropna().tolist())
print("\nAll non-null values in column 4 (TCF/-Plan):")
print(df.iloc[:, 4].dropna().tolist())
print("\nAll non-null values in column 5 (TCF/-VIN):")
print(df.iloc[:, 5].dropna().tolist())
