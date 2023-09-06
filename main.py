import pandas as pd

# Read the first Excel file
df1 = pd.read_excel('file1.xlsx')

# Read the second Excel file
df2 = pd.read_excel('file2.xlsx')

# Merge the two dataframes based on 'Name' and 'Blood Group' columns
merged_df = pd.merge(df1, df2, on=['Name', 'Blood Group'], how='inner')

# Save the merged dataframe to a new Excel file
merged_df.to_excel('merged_file.xlsx', index=False)