import pandas as pd
xf = pd.ExcelFile('TCF VIN  & Paint Float mapping data-1.xlsx')
df = pd.read_excel(xf, sheet_name='Model Wise Float', header=None)
print("Model Wise Float sheet - All rows:")
print(df.to_string())
