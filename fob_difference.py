import pandas as pd

# File paths
old_file = "scenario3_40 tweaks-6.xlsx"
new_file = "scenario3_40 tweaks-6 new.xlsx"

# Columns to read
cols = ["ROW ID #", "Selected Supplier", "Final quote per each FOB Port of Departure (USD)"]

# Read both files (skip first 13 rows)
df_old = pd.read_excel(old_file, skiprows=13, usecols=cols)
df_new = pd.read_excel(new_file, skiprows=13, usecols=cols)

# Merge on Row id + Selected Supplier
merged = df_old.merge(
    df_new,
    on=["ROW ID #", "Selected Supplier"],
    suffixes=("_old", "_new")
)

# Find mismatches
mismatches = merged[
    merged["Final quote per each FOB Port of Departure (USD)_old"] != 
    merged["Final quote per each FOB Port of Departure (USD)_new"]
]

# Select only required columns
result = mismatches[
    ["ROW ID #", "Selected Supplier",
     "Final quote per each FOB Port of Departure (USD)_old",
     "Final quote per each FOB Port of Departure (USD)_new"]
]

# Rename for clarity
result = result.rename(columns={
    "Final quote per each FOB Port of Departure (USD)_old": "Old Value",
    "Final quote per each FOB Port of Departure (USD)_new": "Corrected Value"
})

print(result)

# Optionally save to CSV/Excel
result.to_csv("mismatches.csv", index=False)
