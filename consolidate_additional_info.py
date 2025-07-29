import os
import pandas as pd
from collections import defaultdict

# Paths
ALL_CLEANED_FOLDERS = ["bidsheet_brass", "bidsheet_steel", "bidsheet_other_metal"]
BASE_CLEANED_DIR = "./cleaned_files"
OUTPUT_PATH = "./new/Consolidated Additional Information.xlsx"

# Constants
ROW_KEY_COLUMNS = ["ROW ID #", "Division", 'Part #', "Item Description"]
ADDITIONAL_INFO_COLUMN = "Additional information (please use this column only if absolutely necessary)"

# Data holders
row_key_to_common_data = {}  # row_key -> {common fields}
row_key_to_supplier_values = defaultdict(dict)  # row_key -> {column_header: value}
all_headers = set()

# Process all folders
for folder in ALL_CLEANED_FOLDERS:
    folder_path = os.path.join(BASE_CLEANED_DIR, folder)
    if not os.path.exists(folder_path):
        print(f"[!] Missing folder: {folder_path}")
        continue

    for file_name in os.listdir(folder_path):
        if not file_name.lower().endswith((".xlsx", ".xls")):
            continue

        file_path = os.path.join(folder_path, file_name)
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]

            if ADDITIONAL_INFO_COLUMN not in df.columns:
                continue

            if not all(col in df.columns for col in ROW_KEY_COLUMNS):
                continue

            # Identify supplier and round
            round_tag = "R2" if "_r2" in file_name.lower() else "R1"
            supplier_name = (
                file_name.replace("_r2_cleaned.xlsx", "")
                         .replace("_cleaned.xlsx", "")
                         .replace(".xlsx", "")
                         .replace(".xls", "")
                         .replace("R2", "")
                         .strip()
            )

            col_header = f"{supplier_name.split('--')[-1]} - {round_tag} - {ADDITIONAL_INFO_COLUMN}"
            all_headers.add(col_header)

            for _, row in df.iterrows():
                row_key_parts = []
                common_data = {}
                for col in ROW_KEY_COLUMNS:
                    val = row.get(col, "")
                    clean_val = str(val).strip() if pd.notna(val) else ""
                    row_key_parts.append(clean_val)
                    common_data[col] = clean_val
                row_key = "|".join(row_key_parts)

                if row_key not in row_key_to_common_data:
                    row_key_to_common_data[row_key] = common_data

                value = row.get(ADDITIONAL_INFO_COLUMN, "")
                clean_value = str(value).strip() if pd.notna(value) else "-"
                if clean_value.lower() in ["", "nan"]:
                    clean_value = "-"

                row_key_to_supplier_values[row_key][col_header] = clean_value

            print(f"[✓] Processed: {file_name}")

        except Exception as e:
            print(f"[✗] Failed to read {file_name}: {e}")

# Final headers with supplier columns sorted
sorted_supplier_headers = sorted(all_headers, key=lambda x: x.lower())
final_headers = ROW_KEY_COLUMNS + sorted_supplier_headers

# Build final rows
final_rows = []
for row_key in sorted(row_key_to_common_data):
    base = row_key_to_common_data[row_key].copy()
    for supplier_col in sorted_supplier_headers:
        base[supplier_col] = row_key_to_supplier_values[row_key].get(supplier_col, "-")
    final_rows.append(base)

# Export to Excel with columns in desired order
final_df = pd.DataFrame(final_rows)[final_headers]
os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
final_df.to_excel(OUTPUT_PATH, index=False)

print(f"\n✅ Saved consolidated additional info → {OUTPUT_PATH}")
print(f"📌 Total rows: {len(final_df)}  |  Total suppliers: {len(sorted_supplier_headers)}")
