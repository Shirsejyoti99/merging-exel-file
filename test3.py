import pandas as pd
sheet1 = pd.read_excel('sheet1.xlsx')
sheet2 = pd.read_excel('sheet2.xlsx')
merged_sheet = pd.merge(sheet1, sheet2, on='common_column')
merged_sheet.to_excel('merged_sheet.xlsx', index=False)