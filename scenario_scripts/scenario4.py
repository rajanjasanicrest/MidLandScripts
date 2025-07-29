import pandas as pd
from collections import Counter, defaultdict
from tqdm import tqdm
import time

# --- Start timer ---
start_time = time.time()

# --- Constants ---
TOTAL_COST = 49783721.6480354
THRESHOLD_COST = TOTAL_COST * 0.8  # 80%
VOLUME_COL = "Annual Volume (per UOM)"
INCUMBENT_COL = "Normalized incumbent supplier"
FINAL_MIN_SUPPLIER_COL = "Final Minimum Bid Landed Supplier"
R2_LANDED_COST_SUFFIX = " - R2 - Total landed cost per UOM (USD)"

# --- Load files ---
input_path = "new/Bidsheet Master Consolidate Landed2.csv"
output_reference_file_path = "new/outout-reference.csv"

print("Reading:", input_path)
df = pd.read_csv(input_path)
output_reference_df = pd.read_csv(output_reference_file_path)
print(f"✅ Loaded {len(df)} rows\n")

# --- Step 1: Determine Top 2 Suppliers per Product Group ---
group_supplier_freq = defaultdict(Counter)
incumbent_suppliers = df[INCUMBENT_COL].unique()

print("Counting minimum bid supplier frequency by Product Group...")
for _, row in df.iterrows():
    group = row.get("Product Group")
    supplier = row.get(FINAL_MIN_SUPPLIER_COL)
    if pd.notna(group) and pd.notna(supplier):
        group_supplier_freq[group][supplier] += 1

top2_suppliers_by_group = {
    group: [s for s, _ in counter.most_common(2)]
    for group, counter in group_supplier_freq.items()
}

# --- Step 2: Assign suppliers ---
output_data = []
cumulative_cost = 0
total_fob_savings_usd = 0
total_cost_not_awarded = 0
net_new_supplier_count = 0
parts_where_no_bids = 0
total_landed_cost_incumbent = 0
total_landed_cost_completely_new_suppliers = 0
total_landed_cost_new_suppliers = 0
total_landed_savings_usd = 0
incumbent_retained = 0
new_supplier_count = 0
net_new_supplier_list = set()
unique_suppliers = set()

incumbent_suppliers_list = df["Normalized incumbent supplier"].unique()

print("\nAssigning suppliers with threshold logic...\n")
for _, row in tqdm(df.iterrows(), total=len(df), desc="Finalizing"):

    if row.get("Valid Supplier", 0) == 0:
        selected_supplier = "-"
        reason = "No valid suppliers"
        parts_where_no_bids+=1
        total_cost_not_awarded+=row.get("Landed Extended Cost USD")
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
    
    group = row.get("Product Group")
    incumbent = row.get(INCUMBENT_COL)
    volume = pd.to_numeric(row.get(VOLUME_COL), errors="coerce")
    valid_supplier_count = row.get("Valid Supplier", 0)
    selected_supplier = "-"
    reason = "No valid assignment"

    if valid_supplier_count == 0 or pd.isna(volume) or volume <= 0:
        reason = "No valid suppliers"
    else:
        top_suppliers = top2_suppliers_by_group.get(group, [])
        per_supplier_costs = {}

        # Collect cost for top 2 suppliers (if they bid)
        for supplier in top_suppliers:
            col = f"{supplier}{R2_LANDED_COST_SUFFIX}"
            cost = pd.to_numeric(row.get(col), errors="coerce")
            if pd.notna(cost) and cost > 0:
                per_supplier_costs[supplier] = cost

        if cumulative_cost < THRESHOLD_COST and per_supplier_costs:
            # Assign to the cheaper of the top 2 (if both exist), else fallback
            best_supplier = min(per_supplier_costs, key=per_supplier_costs.get)
            selected_supplier = best_supplier
            extended_cost = per_supplier_costs[best_supplier] * volume
            cumulative_cost += extended_cost
            reason = "Assigned within 80% threshold to top Product Group supplier"
        else:
            # Threshold exceeded or no valid top-2 bid: fallback
            fallback_supplier = row.get(FINAL_MIN_SUPPLIER_COL)
            if pd.notna(fallback_supplier):
                col = f"{fallback_supplier}{R2_LANDED_COST_SUFFIX}"
                cost = pd.to_numeric(row.get(col), errors="coerce")
                if pd.notna(cost) and cost > 0:
                    selected_supplier = fallback_supplier
                    reason = "Threshold exceeded"
                else:
                    selected_supplier = "-"
                    reason = "No valid fallback supplier bid"
            else:
                selected_supplier = "-"
                reason = "No Final Minimum Bid Landed Supplier found"

    # --- Reference supplier data ---
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]

    def get_ref_value(col):
        return ref_row[col].values[0] if not ref_row.empty else "-"
    
    try:
        fob_savings_usd = float(row.get(f"{selected_supplier} - Final USD savings vs baseline", 0)) if pd.notna(row.get(f"{selected_supplier} - Final USD savings vs baseline",  0)) else 0
    except (ValueError, TypeError):
        fob_savings_usd = 0
    
    total_fob_savings_usd += fob_savings_usd 
        
    try:
        lcs = float(row.get(f"{selected_supplier} - Final Landed USD savings vs baseline", 0)) if pd.notna(row.get(f"{selected_supplier} - Final Landed USD savings vs baseline", 0)) else 0
    except (ValueError, TypeError): 
        lcs = 0

    total_landed_savings_usd  += lcs

    if selected_supplier == incumbent:
        total_landed_cost_incumbent += row.get(VOLUME_COL) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)")
        incumbent_retained += 1
    else:
        if selected_supplier in incumbent_suppliers:
            new_supplier_count += 1
            total_landed_cost_new_suppliers += row.get(VOLUME_COL, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        else:
            net_new_supplier_list.add(selected_supplier)
            net_new_supplier_count += 1
            total_landed_cost_completely_new_suppliers += row.get(VOLUME_COL, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)


    if selected_supplier != "-": unique_suppliers.add(selected_supplier) 

    unique_suppliers.add(selected_supplier)

    # --- Output row ---
    output_row = {
        "ROW ID #": row.get("ROW ID #"),
        "Division": row.get("Division"),
        "Part #": row.get("Part #"),
        "Item Description": row.get("Item Description"),
        "Product Group": group,
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "FOB Savings %": row.get(f"{selected_supplier} - Final % savings vs baseline", "-"),
        "FOB Savings USD": row.get(f"{selected_supplier} - Final USD savings vs baseline", "-"),
        "Landed Cost Savings %": row.get(f"{selected_supplier} - Final Landed % savings vs baseline", "-"),
        "Landed Cost Savings USD": row.get(f"{selected_supplier} - Final Landed USD savings vs baseline", "-"),
        "Reason": reason,
        "Landed Extended Cost USD": row[f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"] * row.get(VOLUME_COL, 0),
        "Is Totally New Supplier": "Yes" if (selected_supplier not in incumbent_suppliers and selected_supplier != '-') else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
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
output_file = "Scenario_4_Output.xlsx"

# Write to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario - 4"
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
print(f"\nDone. Output written to '{output_file}.xlsx'")
print(f"⏱Time taken: {elapsed_time:.2f} seconds")

