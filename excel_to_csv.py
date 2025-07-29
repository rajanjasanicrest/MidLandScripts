import pandas as pd

# === CONFIGURATION ===
excel_file_path = "new/Bidsheet Master Consolidate Landed3.xlsx"
sheet_name = "Sheet1"
csv_output_path = "new/Bidsheet Master Consolidate Landed3.csv"

# === CONVERSION ===
print(f"Reading Excel file: {excel_file_path}")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
print(f"Saving as CSV: {csv_output_path}")
df.to_csv(csv_output_path, index=False)
print("Conversion complete.")

# file_paths =[
#     "constants/Uncompetitive supplier behavior table 071025.xlsx",
#     "constants/Standard leadtime table 071025.xlsx",
#     "constants/Retail Packaging table 071225.xlsx",
#     "constants/Payment term table 071025.xlsx",
#     "constants/New product introduction table 071025.xlsx",
#     "constants/Long term commitment rebate table 071225.xlsx",
# ]

# custom_names = [
#     "Uncompetitive supplier behavior",
#     "Standard leadtime - days PO-shipment POL",
#     "Retail Packaging",
#     "Payment term - days and discounts",
#     "New product introduction",
#     "Long term commitment rebate"
# ]

# final_df = None
# for file_path, custom_col in zip(file_paths, custom_names):
#     df = pd.read_excel(file_path, skiprows=2)

#     # Rename the Output column to the custom name
#     df = df.rename(columns={"Output": custom_col})
    
#     if final_df is None:
#         final_df = df
#     else:
#         # Merge on Reference column
#         final_df = pd.merge(final_df, df, on="Reference", how="outer")

# # --- Output ---
# # Save to a new Excel file
# final_df.to_csv("new/outout-reference.csv", index=False)
