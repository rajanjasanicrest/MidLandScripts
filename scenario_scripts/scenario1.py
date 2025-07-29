import pandas as pd
from tqdm import tqdm
import time

# --- Start timing ---
start_time = time.time()

# --- Constants ---
BASE_PERCENTAGE = 5

# --- Load Excel ---
input_path = "new/Bidsheet Master Consolidate Landed2.csv"
print('reading file:', input_path)
df = pd.read_csv(input_path)
print(f"✅ Loaded {len(df)} rows from '{input_path}'\n")

output_reference_file_path = 'new/outout-reference.csv'
output_reference_df = pd.read_csv(output_reference_file_path)

# --- Define column names ---
incumbent_col = "Normalized incumbent supplier"
lowest_bid_col = "Final Minimum Bid Landed Supplier"
savings_pct_col = "As Is Final Landed %"
savings_usd_col = "As Is Final Landed USD"
volume_col = "Annual Volume (per UOM)"

# --- Final output columns ---
columns_to_keep = [
    "ROW ID #", "Division", "Part #", "Item Description",
    "Product Group", "Part Family"
]
incumbent_suppliers = df[incumbent_col].unique()

# --- Prepare output ---
output_data = []
fob_savings_pct = []
fob_savings_usd = []
total_fob_savings_usd = 0
total_landed_savings_usd = 0
incumbent_retained = 0
total_landed_cost_incumbent=0
total_landed_cost_new_suppliers=0
net_new_supplier_count = 0
new_supplier_count = 0
total_cost_not_awarded = 0
net_new_supplier_count = 0
parts_where_no_bids = 0
total_landed_cost_completely_new_suppliers = 0
unique_suppliers = set()

print("Processing rows with Scenario 1 logic...\n")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing"):

    incumbent = row.get(incumbent_col)
    lowest_bidder = row.get(lowest_bid_col)

    if row.get("Valid Supplier", 0) == 0:
        total_cost_not_awarded += row.get('Landed Extended Cost USD', 0)
        selected_supplier = "No valid suppliers"
        parts_where_no_bids+=1
        reason = "No valid suppliers"
        output_row = {
            "ROW ID #": row.get("ROW ID #"),
            "Division": row.get("Division"),
            "Part #": row.get("Part #"),
            "Item Description": row.get("Item Description"),
            "Product Group": row.get("Product Group"),
            "Part Family": row.get("Part Family"),
            "Incumbent Supplier": incumbent,
            "Selected Supplier": selected_supplier,
            "FOB Savings %": 0,
            "FOB Savings USD": 0,
            "Landed Cost Savings %": 0,
            "Landed Cost Savings USD": 0,
            "Reason": reason,
            "Landed Extended Cost USD": 0,
            "Is Totally New Supplier": "-",
            "Part Switched": "-",
            "Standard leadtime - days PO-shipment POL": '-',
            "Retail Packaging": '-',
            "Payment term - days and discounts": '-',
            "New product introduction": '-',
            "Long term commitment rebate": '-',
            "Uncompetitive supplier behavior": '-',
        }
        output_data.append(output_row)
        continue
    
    savings_pct = row.get(savings_pct_col)
    try:
        savings_pct = float(savings_pct) if pd.notna(savings_pct) else 0
    except (ValueError, TypeError):
        savings_pct = 0

    savings_usd = row.get(savings_usd_col)
    try:
        savings_usd = float(savings_usd) if pd.notna(savings_usd) else 0
    except (ValueError, TypeError):
        savings_usd = 0

    if pd.notna(savings_pct) and savings_pct >= BASE_PERCENTAGE:
        bid_check_col = f"{incumbent} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
        bid_val = row.get(bid_check_col)

        if bid_check_col in df.columns and pd.notna(bid_val) and bid_val > 0:
            selected_supplier = incumbent
            reason = f"Incumbent retained (≥ {BASE_PERCENTAGE}% savings)"
        else:
            selected_supplier = lowest_bidder
            reason = f"Incumbent not available, switched to lowest bidder"

    elif incumbent == lowest_bidder:
        selected_supplier = incumbent
        reason = "Incumbent retained (same as lowest bidder)"
    else:
        selected_supplier = lowest_bidder
        reason = f"Switched to lowest bidder (< {BASE_PERCENTAGE}% savings)"
        savings_pct = row.get(f"{selected_supplier} - Final % savings vs baseline", "-")
        savings_usd = row.get(f"{selected_supplier} - Final USD savings vs baseline", "-")
        try:
            savings_pct = float(savings_pct) if pd.notna(savings_pct) else 0
        except (ValueError, TypeError):
            savings_pct = 0

        try:
            savings_usd = float(savings_usd) if pd.notna(savings_usd) else 0
        except (ValueError, TypeError):
            savings_usd = 0

    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
    if selected_supplier != "No valid suppliers":
        landed_savings_pct = row.get(landed_pct_col, "-")
        landed_savings_usd = row.get(landed_usd_col, "-")

    total_fob_savings_usd += savings_usd 
    try:
        lcs = float(landed_savings_usd) if pd.notna(landed_savings_usd) else 0
    except (ValueError, TypeError): 
        lcs = 0

    total_landed_savings_usd  += lcs

    if selected_supplier == incumbent:
        total_landed_cost_incumbent += row.get(volume_col) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)")
        incumbent_retained += 1
    else:
        if selected_supplier in incumbent_suppliers:
            new_supplier_count += 1
            total_landed_cost_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        else:
            net_new_supplier_count += 1
            total_landed_cost_completely_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)

    if selected_supplier != "-": unique_suppliers.add(selected_supplier) 

    # print(selected_supplier, reason)
    output_row = {
        "ROW ID #": row.get("ROW ID #"),
        "Division": row.get("Division"),
        "Part #": row.get("Part #"),
        "Item Description": row.get("Item Description"),
        "Product Group": row.get("Product Group"),
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "FOB Savings %": savings_pct,
        "FOB Savings USD": savings_usd,
        "Landed Cost Savings %": landed_savings_pct,
        "Landed Cost Savings USD": landed_savings_usd,
        "Reason": reason,
        "Landed Extended Cost USD": row[f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"] * row.get(volume_col, 0),
        "Is Totally New Supplier": "Yes" if selected_supplier not in incumbent_suppliers else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
        "Standard leadtime - days PO-shipment POL": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "Standard leadtime - days PO-shipment POL"].values[0] if selected_supplier != 'No valid suppliers' else '-',
        "Retail Packaging": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "Retail Packaging"].values[0] if selected_supplier != 'No valid suppliers' else '-',
        "Payment term - days and discounts": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "Payment term - days and discounts"].values[0] if selected_supplier != 'No valid suppliers' else '-',
        "New product introduction": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "New product introduction"].values[0] if selected_supplier != 'No valid suppliers' else '-',
        "Long term commitment rebate": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "Long term commitment rebate"].values[0] if selected_supplier != 'No valid suppliers' else '-',
        "Uncompetitive supplier behavior": output_reference_df.loc[output_reference_df['Reference'] == selected_supplier, "Uncompetitive supplier behavior"].values[0] if selected_supplier != 'No valid suppliers' else '-',
    }
    output_data.append(output_row)

# Calculate redundant suppliers per product family
from collections import defaultdict

# Step 1: Map product family to all selected suppliers
family_supplier_map = defaultdict(set)
for row in output_data:
    pf = row["Product Group"]
    supplier = row["Selected Supplier"]
    if supplier != "-":
        family_supplier_map[pf].add(supplier)

# Step 2: Calculate the count of unique suppliers per family
redundancy_count_map = {pf: len(suppliers) for pf, suppliers in family_supplier_map.items()}

# Step 3: Add redundancy value to each output row
for row in output_data:
    pf = row["Product Group"]
    if pf == "No group available":
        row["Redundant Suppliers per Product Group"] = 1
    else:
        row["Redundant Suppliers per Product Group"] = redundancy_count_map.get(pf, 0)

# Step 4: Convert to DataFrame
output_df = pd.DataFrame(output_data)

# Step 5: Move column to index 14
if "Redundant Suppliers per Product Group" in output_df.columns:
    redundancy_col = output_df.pop("Redundant Suppliers per Product Group")
    output_df.insert(14, "Redundant Suppliers per Product Group", redundancy_col)

incumbent_rows = output_df[output_df["Selected Supplier"] == output_df["Incumbent Supplier"]]
incumbent_suppliers_unique = set(incumbent_rows["Selected Supplier"].dropna().unique())
if '-' in incumbent_suppliers_unique: incumbent_suppliers_unique.remove('-')

summary_data = [
    # ["Total FOB Savings USD", total_fob_savings_usd],
    ["Total Landed Cost Savings USD", total_landed_savings_usd],
    ["Total Cost of no valid suppliers", total_cost_not_awarded],
    ["Total Landed Cost where Incumbent Suppliers Retained", total_landed_cost_incumbent],
    ["Total Landed Cost where bid is awarded to Net New Suppliers", total_landed_cost_completely_new_suppliers],
    ["Total Landed Cost where bid is awarded to New Suppliers", total_landed_cost_new_suppliers],
    ["Total parts where Incumbent Suppliers Retained", incumbent_retained],
    ["Total parts where bid is awarded to New Suppliers", new_supplier_count],
    ["Total parts where bid is awarded to Net New Suppliers", net_new_supplier_count],
    ["Parts not awarded to any supplier", parts_where_no_bids],
    ["", ""],
    
    ["Totally New Suppliers", len(unique_suppliers)-len(incumbent_suppliers_unique)],
    ["Total Unique Suppliers", len(unique_suppliers)],

]
# Output file
output_file = "Scenario_1_Output.xlsx"

# Write to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario - 1: Incumbent Retained if they provide savings of at least {BASE_PERCENTAGE}%, else move to new supplier"
    header_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
    worksheet.merge_range(0, 0, 0, len(list(output_data[0].keys()))-1, scenario_header, header_format)

    # Define formats
    bold_format = workbook.add_format({'bold': True})
    usd_format = workbook.add_format({'num_format': '$000,00.00'})

    # Write summary into separate logical blocks
    summary_row = 2

    # Grouped Indiced
    cost_metrics = summary_data[:5]
    supplier_metrics = summary_data[5:15]
    volume_metrics = summary_data[15:]

    # Write the summary manually at top (row 0 onward)
    for row_idx, row in enumerate(cost_metrics):
        worksheet.write(summary_row + row_idx, 0, row[0], bold_format)
        worksheet.write(summary_row + row_idx, 1, row[1], usd_format)

    for i, item in enumerate(supplier_metrics):
        worksheet.write(summary_row + i, 3, item[0], bold_format)
        worksheet.write(summary_row + i, 4, item[1])

    # --- Write total evaluated cost row ---
    total_label_row = summary_row + max(len(cost_metrics), len(supplier_metrics), len(volume_metrics)) + 1
    worksheet.write(total_label_row, 0, "Total Landed Cost Evaluated", bold_format)

    # Formula for summing cost values (adjust B3:B8 if more/less than 6 rows of cost)
    worksheet.write_formula(total_label_row, 1, '=SUM(B3:B8)', usd_format)

    # Write the DataFrame starting at row 8 (leaving some space)
    output_df.to_excel(writer, sheet_name="Sheet1", startrow=len(summary_data) + 2, index=False)

# --- Print total time taken ---
elapsed_time = time.time() - start_time
print(f"\nDone. Output written to 'Scenario_1_Output.xlsx'")
print(f"⏱Time taken: {elapsed_time:.2f} seconds")
