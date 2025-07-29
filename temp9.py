import pandas as pd

# Read sheet 2 (index 1) from the Excel file
df = pd.read_excel('new/discount_rebate_consolidate.xlsx', sheet_name=1)

# Save the DataFrame to CSV
df.to_csv('new/discount.csv', index=False)