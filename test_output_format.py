import pandas as pd
import numpy as np

xf = pd.ExcelFile('TCF VIN  & Paint Float mapping data-1.xlsx')
sheets_to_inspect = ['Model Wise Float', 'Engine Summary', 'Wiring Summary', 'Cockpit Summary']

for sheet in sheets_to_inspect:
    print(f"\n{'='*70}")
    print(f"SHEET: {sheet}")
    df = pd.read_excel(xf, sheet_name=sheet, header=None)
    print(f"Shape: {df.shape}")
    print(f"\nAll rows (raw):")
    print(df.to_string())
    print()
