import pandas as pd
import re
import numpy as np
 
# ========== FILE INPUT ==========
file_info = [
    ("consolidate/bidsheet_steel_outlier_consolidate.xlsx", "steel"),
    ("consolidate/bidsheet_brass_outlier_consolidate.xlsx", "brass"),
    ("consolidate/bidsheet_other_metal_outlier_consolidate.xlsx", "other metal")
]
 
dfs = []
 
for file_path, metal_type in file_info:
    df = pd.read_excel(file_path, header=[0, 1])
   
    # Ensure unique level-1 headers to avoid merging by Excel
    new_cols = []
    counter = {}
 
    for col in df.columns:
        lvl1, lvl2 = col
        if lvl1.startswith("Unnamed") or pd.isna(lvl1):
            new_cols.append(("", lvl2))
        else:
            base = lvl1.strip()
            counter[base] = counter.get(base, 0) + 1
            unique_lvl1 = f"{base}_{counter[base]}"  # temp suffix
            new_cols.append((unique_lvl1, lvl2))
 
    df.columns = pd.MultiIndex.from_tuples(new_cols)
 
    # Add 'type' column after 5th column
    df.insert(5, ('', 'type'), metal_type)
 
    dfs.append(df)
 
# ========== COMBINE DATA ==========
combined_df = pd.concat(dfs, axis=0, ignore_index=True)
 
# Column arrangement: 5 initial + type, then middle, then last 3
# Separate columns into: first_6, middle, last_3
first_6 = dfs[0].columns[:5].tolist() + [('', 'type')]
last_3 = dfs[0].columns[-3:].tolist()
 
# Get the rest (middle columns)
middle = [col for col in combined_df.columns if col not in first_6 and col not in last_3]
 
# Sort middle columns by level-1 (company name)
middle_sorted = sorted(middle, key=lambda x: (re.sub(r"_\d+$", "", x[0])))
 
# Reorder DataFrame columns
combined_df = combined_df.loc[:, first_6 + middle_sorted + last_3]
 
# ========== CLEAN HEADER ==========
cleaned_cols = []
for lvl1, lvl2 in combined_df.columns:
    if lvl1 == "":
        cleaned_cols.append(("", lvl2))
    else:
        cleaned_lvl1 = re.sub(r"_\d+$", "", lvl1)
        cleaned_cols.append((cleaned_lvl1, lvl2))
 
combined_df.columns = pd.MultiIndex.from_tuples(cleaned_cols)
 
# ========== REMOVE EMPTY "ADDITIONAL INFORMATION" COLUMNS ==========
columns_to_remove = []
 
for col in combined_df.columns:
    lvl1, lvl2 = col
    # Check if column header contains "Additional information" text
    if ("additional information" in str(lvl2).lower() and
        "please use this column only if absolutely necessary" in str(lvl2).lower()):
       
        # Check if the column has any non-null, non-empty data
        has_data = False
        for value in combined_df[col]:
            if pd.notna(value) and str(value).strip() != "":
                has_data = True
                break
       
        # If no data found, mark for removal
        if not has_data:
            columns_to_remove.append(col)
            print(f"Removing empty 'Additional information' column: {col}")
 
# Remove the empty additional information columns
if columns_to_remove:
    combined_df = combined_df.drop(columns=columns_to_remove)
    print(f"Removed {len(columns_to_remove)} empty 'Additional information' columns")
else:
    print("No empty 'Additional information' columns found to remove")
 
# ========== FORMAT FLOAT VALUES TO 4 DECIMAL PLACES AND EXCLUDE ZEROS (FROM 5TH COLUMN) ==========
def format_value(value):
    """Format numeric values to 4 decimal places, exclude zeros, keep non-numeric as is"""
    if pd.isna(value):
        return ""
    elif isinstance(value, (int, float, np.number)):
        if value == 0:
            return "0"  # Exclude zeros
        else:
            return round(float(value), 4)  # Convert to 4 decimal places
    else:
        return value  # Keep strings and other types as is
 
# Apply formatting only from 5th column onwards (index 4 and beyond)
for col_idx, col in enumerate(combined_df.columns):
    if col_idx >= 4:  # From 5th column (index 4) onwards
        combined_df[col] = combined_df[col].apply(format_value)
 
# ========== EXPORT TO EXCEL WITH STYLING ==========
with pd.ExcelWriter("combined_bidsheet_outlier_2.xlsx", engine="xlsxwriter") as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet
 
    # Create fill formats
    yellow_fill = workbook.add_format({'bg_color': '#FFFF99'})
    green_fill = workbook.add_format({'bg_color': '#90EE90'})
    gray_fill = workbook.add_format({'bg_color': '#D3D3D3'})
    white_fill = workbook.add_format({'bg_color': '#FFFFFF'})
   
    # Create number format for 4 decimal places
    decimal_format = workbook.add_format({'num_format': '0.0000'})
    decimal_yellow = workbook.add_format({'bg_color': '#FFFF99', 'num_format': '0.0000'})
    decimal_green = workbook.add_format({'bg_color': '#90EE90', 'num_format': '0.0000'})
    decimal_gray = workbook.add_format({'bg_color': '#D3D3D3', 'num_format': '0.0000'})
    decimal_white = workbook.add_format({'bg_color': '#FFFFFF', 'num_format': '0.0000'})
 
    # Flatten headers (update after column removal)
    cleaned_cols = combined_df.columns.tolist()
    level_1_headers = [col[0] for col in cleaned_cols]
    level_2_headers = [col[1] for col in cleaned_cols]
 
    # Write headers with alternating fills for middle company blocks
    current_company = None
    fill_toggle = True
    for col_idx, (lvl1, lvl2) in enumerate(cleaned_cols):
        if 5 <= col_idx < len(cleaned_cols) - 3 and lvl1:
            if lvl1 != current_company:
                current_company = lvl1
                fill_toggle = not fill_toggle
            fill_format = gray_fill if fill_toggle else white_fill
            worksheet.write(0, col_idx, f'{lvl1}-{lvl2}', fill_format)
        else:
            worksheet.write(0, col_idx, f'{lvl2}')
 
    # Write data rows with proper formatting
    for row_idx, row in enumerate(combined_df.values, start=2):
        for col_idx, value in enumerate(row):
            # Apply decimal formatting only from 5th column onwards (index 4+)
            if col_idx >= 4 and isinstance(value, (int, float, np.number)) and not pd.isna(value):
                # Apply number formatting for numeric values from 5th column onwards
                if col_idx == 4:
                    worksheet.write_number(row_idx, col_idx, value, decimal_green)
                elif col_idx >= len(cleaned_cols) - 3:
                    worksheet.write_number(row_idx, col_idx, value, decimal_yellow)
                elif 5 <= col_idx < len(cleaned_cols) - 3:
                    # For middle company columns, determine fill color
                    current_company = None
                    fill_toggle = True
                    for c_idx, (lvl1, lvl2) in enumerate(cleaned_cols[:col_idx+1]):
                        if 5 <= c_idx < len(cleaned_cols) - 3 and lvl1:
                            if lvl1 != current_company:
                                current_company = lvl1
                                fill_toggle = not fill_toggle
                    format_to_use = decimal_gray if fill_toggle else decimal_white
                    worksheet.write_number(row_idx, col_idx, value, format_to_use)
                else:
                    worksheet.write_number(row_idx, col_idx, value, decimal_format)
            else:
                # Write values from first 4 columns or non-numeric values as text
                cell_value = "" if pd.isna(value) or value == "" else str(value)
                # Apply background color formatting for specific columns even for text
                if col_idx == 4:
                    worksheet.write(row_idx, col_idx, cell_value, green_fill)
                elif col_idx >= len(cleaned_cols) - 3:
                    worksheet.write(row_idx, col_idx, cell_value, yellow_fill)
                else:
                    worksheet.write(row_idx, col_idx, cell_value)
 
    # Set column widths for better visibility
    for col_idx in range(len(cleaned_cols)):    
        worksheet.set_column(col_idx, col_idx, 12)
 
print("Processing complete! File saved as 'combined_bidsheet_outlier_2.xlsx'")
print("- Float values from 7th column onwards formatted to 4 decimal places")
print("- Zero values from 7th column onwards excluded (shown as empty cells)")
print("- First 6 columns retain original formatting")
print("- Empty 'Additional information' columns removed")
print("- Styling maintained for different column groups")