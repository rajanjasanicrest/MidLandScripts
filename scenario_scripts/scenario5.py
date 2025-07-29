"""
📦 SCENARIO 5 WORKFLOW: GROW-THEN-LOCK SUPPLIER POOL (MAX 20)

1. Start with an empty supplier pool.
2. For each part:
   a. If supplier pool size < 20:
       - If any supplier in current pool has valid bid → assign to lowest among them.
       - Else → assign to Final Minimum Bid Landed Supplier & add to pool.
   b. If pool has 20 suppliers:
       - Try assigning to lowest valid bidder from the pool.
       - If none of the 20 suppliers bid on the part:
           - Assign to global lowest bidder (from all suppliers with a valid bid).
3. Never assign to incumbent (even in fallback).
4. Cap supplier pool strictly to 20 — no replacement logic.
5. Attach savings columns & reference data like previous scenarios.
"""

import pandas as pd
from tqdm import tqdm
import time

# --- Start timer ---
start_time = time.time()

# --- File paths ---
input_path = "new/Bidsheet Master Consolidate Landed2.csv"
output_reference_file_path = "new/outout-reference.csv"

# --- Constants ---
incumbent_col = "Normalized incumbent supplier"
valid_supplier_col = "Valid Supplier"
volume_col = 'Annual Volume (per UOM)'

# --- Load data ---
print("Reading:", input_path)
df = pd.read_csv(input_path)
output_reference_df = pd.read_csv(output_reference_file_path)
print(f"✅ Loaded {len(df)} rows\n")
incumbent_suppliers = df[incumbent_col].unique()

# --- Identify landed cost columns ---
r2_cost_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_cost_cols]

# --- Track supplier pool ---
supplier_pool = []
supplier_pool_set = set()

# --- Output data ---
output_data = []
total_cost_not_awarded = 0
total_landed_cost_incumbent = 0
total_landed_cost_completely_new_suppliers = 0
total_landed_cost_new_suppliers = 0
fallback_counter = 0
parts_where_no_bids = 0
cumulative_cost = 0
net_new_supplier_count = 0
total_fob_savings_usd = 0
total_landed_savings_usd = 0
incumbent_retained = 0
new_supplier_count = 0
unique_suppliers = set()
net_new_supplier_list = set()

incumbent_suppliers_list = df["Normalized incumbent supplier"].unique()


import numpy as np
def get_sorting_value(row):
    if row["Valid Supplier"] == 0:
        return np.nan
    supplier = row["Final Minimum Bid Landed Supplier"] or row["Normalized incumbent supplier"]
    savings_col = f"{supplier} - Final Landed USD savings vs baseline"
    return row.get(savings_col, np.nan)

# Compute sorting column
df["sorting_savings"] = df.apply(get_sorting_value, axis=1)

# Drop rows where sorting value is NaN
df_sorted = df[df["sorting_savings"].notna()].copy()

# Sort in descending order (reverse sort)
df_sorted = df_sorted.sort_values(by="sorting_savings", ascending=False)

for idx, row in tqdm(df.iterrows(), total=len(df), desc="Assigning"):

    if row.get("Valid Supplier", 0) == 0:
        selected_supplier = "-"
        reason = "No valid suppliers"
        parts_where_no_bids+=1
        total_cost_not_awarded+=row.get('Landed Extended Cost USD')
        output_row = {
            "ROW ID #": row.get("ROW ID #"),
            "Division": row.get("Division"),
            "Part #": row.get("Part #"),
            "Item Description": row.get("Item Description"),
            "Product Group": row.get("Product Group"),
            "Part Family": row.get("Part Family"),
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
    
    else:
        final_min_supplier = row.get("Final Minimum Bid Landed Supplier") 
        # Try to use pool
        candidate_suppliers = []

        if len(supplier_pool) < 20:
            # Check if pool has any bidder
            for supplier in supplier_pool:
                col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
                cost = row.get(col)
                if pd.notna(cost) and cost > 0:
                    candidate_suppliers.append((supplier, cost))

            if candidate_suppliers:
                selected_supplier = min(candidate_suppliers, key=lambda x: x[1])[0]
                reason = "Selected from partial pool"
            else:
                selected_supplier = final_min_supplier
                if pd.notna(selected_supplier) and selected_supplier not in supplier_pool_set:
                    supplier_pool.append(selected_supplier)
                    supplier_pool_set.add(selected_supplier)
                reason = "Added to supplier pool (<20)"
        else:
            # Pool is full — use lowest from within pool
            for supplier in supplier_pool:
                col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
                cost = row.get(col)
                if pd.notna(cost) and cost > 0:
                    candidate_suppliers.append((supplier, cost))

            if candidate_suppliers:
                selected_supplier = min(candidate_suppliers, key=lambda x: x[1])[0]
                reason = "Selected from 20 supplier pool"
            else:
                # Fallback to global lowest
                global_candidates = []
                for supplier in all_suppliers:
                    col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
                    cost = row.get(col)
                    if pd.notna(cost) and cost > 0:
                        global_candidates.append((supplier, cost))
                if global_candidates:
                    selected_supplier = min(global_candidates, key=lambda x: x[1])[0]
                    reason = "Fallback to global lowest bidder"
                    fallback_counter += 1
                else:
                    selected_supplier = "-"
                    reason = "No valid bids from any supplier"

    # --- Savings columns ---
    pct_col = f"{selected_supplier} - Final % savings vs baseline"
    usd_col = f"{selected_supplier} - Final USD savings vs baseline"
    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"

    # --- Reference info ---
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
    def get_ref_value(col):
        return ref_row[col].values[0] if not ref_row.empty else "-"

    try:
        fob_savings_usd = float(row.get(usd_col, 0)) if pd.notna(row.get(usd_col,  0)) else 0
    except (ValueError, TypeError):
        fob_savings_usd = 0
    
    total_fob_savings_usd += fob_savings_usd 
        
    try:
        lcs = float(row.get(landed_usd_col, 0)) if pd.notna(row.get(landed_usd_col, 0)) else 0
    except (ValueError, TypeError): 
        lcs = 0

    total_landed_savings_usd  += lcs

    if selected_supplier == row.get(incumbent_col):
        incumbent_retained += 1
        total_landed_cost_incumbent += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        unique_suppliers.add(selected_supplier)

    else:
        if selected_supplier in incumbent_suppliers:
            new_supplier_count += 1
            total_landed_cost_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        else:
            net_new_supplier_count += 1
            net_new_supplier_list.add(selected_supplier)
            total_landed_cost_completely_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)

        unique_suppliers.add(selected_supplier)
    output_row = {
        "ROW ID #": row.get("ROW ID #"),
        "Division": row.get("Division"),
        "Part #": row.get("Part #"),
        "Item Description": row.get("Item Description"),
        "Product Group": row.get("Product Group"),
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "FOB Savings %": row.get(pct_col, "-"),
        "FOB Savings USD": row.get(usd_col, "-"),
        "Landed Cost Savings %": row.get(landed_pct_col, "-"),
        "Landed Cost Savings USD": row.get(landed_usd_col, "-"),
        "Reason": reason,
        "Landed Extended Cost USD": row[f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"] * row.get(volume_col, 0),
        "Is Totally New Supplier": "Yes" if (selected_supplier not in incumbent_suppliers and selected_supplier != '-') else "No",
        "Part Switched": "Yes" if selected_supplier != row.get('Normalized incumbent supplier') else "No",
        "Standard leadtime - days PO-shipment POL": get_ref_value("Standard leadtime - days PO-shipment POL"),
        "Retail Packaging": get_ref_value("Retail Packaging"),
        "Payment term - days and discounts": get_ref_value("Payment term - days and discounts"),
        "New product introduction": get_ref_value("New product introduction"),
        "Long term commitment rebate": get_ref_value("Long term commitment rebate"),
        "Uncompetitive supplier behavior": get_ref_value("Uncompetitive supplier behavior"),
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


print(unique_suppliers)
print(incumbent_suppliers_unique)

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
    
    ["Totally New Suppliers", len(net_new_supplier_list)],
    ["Total Unique Suppliers", len(unique_suppliers)],
]

# Output file
output_file = "Scenario_5_Output.xlsx"

# Write to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario - 5"
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
print(f"\nDone. Output written to {output_file}")
print(f"⏱Time taken: {elapsed_time:.2f} seconds")

