import pandas as pd, numpy as np, re

file = r'd:\TML PPC Dashboard\TCF VIN  & Paint Float mapping data-1.xlsx'

# TCF1 Cockpit detailed parse
df = pd.read_excel(file, sheet_name='TCF1 Cockpit', header=None)
h = 3  # header row
header = df.iloc[h]
print("Header columns with values:")
for j, v in enumerate(header):
    if pd.notna(v):
        print(f"  col {j}: {v}")

data = df.iloc[h+1:].copy().reset_index(drop=True)
# Part No at col 0, AB12 Part No at col 3, VC No at col 4
count = 0
for idx, row in data.head(10).iterrows():
    part = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
    vc = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ''
    model = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ''
    print(f"  Row {idx}: part={part}, vc={vc}, model={model[:30]}")
    if bool(re.match(r'^\d{7,9}[A-Za-z]?$', vc)):
        count += 1
print(f"\nValid cockpit rows in first 10: {count}")

# Check all
total = 0
for _, row in data.iterrows():
    vc = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ''
    if bool(re.match(r'^\d{7,9}[A-Za-z]?$', vc)):
        total += 1
print(f"Total valid cockpit rows: {total}")

# DPT Plan test
df2 = pd.read_excel(file, sheet_name='TCF1 DPT Plan', header=None)
print(f"\nTCF1 DPT Plan shape: {df2.shape}")
print("Row 0:", [(j, v) for j, v in enumerate(df2.iloc[0]) if pd.notna(v)])
print("Row 1:", [(j, v) for j, v in enumerate(df2.iloc[1]) if pd.notna(v)][:10])

df3 = pd.read_excel(file, sheet_name='TCF2 DPT Plan', header=None)
print(f"\nTCF2 DPT Plan shape: {df3.shape}")
print("Row 0:", [(j, v) for j, v in enumerate(df3.iloc[0]) if pd.notna(v)])
print("Row 1:", [(j, v) for j, v in enumerate(df3.iloc[1]) if pd.notna(v)][:10])
