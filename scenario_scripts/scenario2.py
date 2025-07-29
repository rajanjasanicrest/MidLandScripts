import pandas as pd
from tqdm import tqdm
import time

# --- Start timing ---
start_time = time.time()

# --- Constants ---
input_path = "new/Bidsheet Master Consolidate Landed2.csv"
output_path = "Scenario_2_Output.xlsx"
priority_suppliers = ["Mayank", "Ningbo Huaping"]
split_keywords = ["BLACK AND GALV", "NIPPLES, BLACK STEEL", "NIPPLES, STAINLESS STEEL", "STEEL NIPPLES", "NIPPLES, GALVANIZED STEEL", "SCH 40 AND SCH 80 GROOVE NIPPLES FITTINGS AND NIPPLES", "SCH 40 STAINLESS STEEL NIPPLES - 304SS"]
volume_col = 'Annual Volume (per UOM)'
incumbent_col = "Normalized incumbent supplier"
# --- Load CSV ---
print("Reading file:", input_path)
df = pd.read_csv(input_path)
print(f"✅ Loaded {len(df)} rows from '{input_path}'\n")

output_reference_file_path = 'new/outout-reference.csv'
output_reference_df = pd.read_csv(output_reference_file_path)

# --- Prepare output ---
output_data = []
total_fob_savings_usd = 0
total_landed_savings_usd = 0
incumbent_retained = 0
total_cost_not_awarded = 0
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0 
total_landed_cost_completely_new_suppliers = 0
new_supplier_count = 0
net_new_supplier_count = 0
parts_where_no_bids = 0
new_supplier_count = 0
unique_suppliers = set()

net_new_supplier_list = set()

incumbent_suppliers_list = df[incumbent_col].unique()

# --- Step 1: Enforce supplier diversity in split product groups ---
split_group_awards = {}
incumbent_suppliers = df[incumbent_col].unique()

for group in df['Product Group'].dropna().unique():
    group_upper = group.upper()
    if not any(keyword in group_upper for keyword in split_keywords):
        continue  # Skip non-split groups

    group_rows = df[df['Product Group'].str.upper() == group_upper].copy()
    group_rows = group_rows.sort_values(by="Cherry Pick Landed Final USD", ascending=False)
    awarded_suppliers = set()
    assigned_rows = {}

    for _, row in group_rows.iterrows():
        if row.get("Valid Supplier", 0) == 0:
            row_id = row["ROW ID #"]
            assigned_rows[row_id] = {
                "Selected Supplier": "-",
                "Reason": "No valid suppliers in split group"
            }
            continue

        elif row.get("Valid Supplier") ==1 :
            row_id = row['ROW ID #']
            assigned_rows[row_id] = {
                "Selected Supplier": row.get('Final Minimum Bid Landed Supplier'),
                "Reason": "Only one Supplier"
            }
            awarded_suppliers.add(row.get('Final Minimum Bid Landed Supplier'))
            continue

        row_id = row["ROW ID #"]
        min_supplier = row.get("Final Minimum Bid Landed Supplier")
        second_supplier = row.get("2nd Lowest Bid Landed Supplier")
        incumbent = row.get("Normalized incumbent supplier")

        if not min_supplier:
            assigned_rows[row_id] = {
                "Selected Supplier": "No valid suppliers",
                "Reason": "No valid suppliers in split group"
            }
            continue

        if len(awarded_suppliers) < 2:
            # Pick supplier not yet awarded, fallback to min_supplier
            if min_supplier not in awarded_suppliers:
                selected = min_supplier
            elif second_supplier and second_supplier not in awarded_suppliers:
                selected = second_supplier
            else:
                selected = min_supplier
        else:
            selected = min_supplier

        awarded_suppliers.add(selected)
        reason = f"Rule 2: Diversity enforced in '{group_upper}' — awarded to {selected}"
        assigned_rows[row_id] = {
            "Selected Supplier": selected,
            "Reason": reason
        }

    split_group_awards.update(assigned_rows)


print("Processing rows with Scenario 2 logic...\n")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing"):

    selected_supplier = None
    other_supplier = '-'
    incumbent = row.get("Normalized incumbent supplier")

    if row.get("Valid Supplier", 0) == 0:
        parts_where_no_bids += 1
        reason = "No valid suppliers"
        total_cost_not_awarded += row.get("Landed Extended Cost USD")
        output_row = {
            "ROW ID #": row.get("ROW ID #"),
            "Division": row.get("Division"),
            "Part #": row.get("Part #"),
            "Item Description": row.get("Item Description"),
            "Product Group": row.get("Product Group"),
            "Part Family": row.get("Part Family"),
            "Incumbent Supplier": incumbent,
            "Selected Supplier": '-',
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

    # --- Extract key data ---
    row_id = row.get("ROW ID #")
    part_no = row.get("Part #")
    product_group = str(row.get("Product Group", "")).upper()
    
    valid_supplier_count = row.get("Valid Supplier", 0)
    
    min_bid = row.get("Final Min Bid")
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
    min2_bid = row.get("Final 2nd Lowest Bid")
    min2_supplier = row.get("2nd Lowest Bid Landed Supplier")

    reason = ""
    fob_savings_pct = None
    fob_savings_usd = None
    landed_savings_pct = None
    landed_savings_usd = None

    # --- Rule 2: Product Group Split (New Logic using split_group_awards) ---
    if any(keyword in product_group for keyword in split_keywords):
        assigned = split_group_awards.get(row_id, {})
        selected_supplier = assigned.get("Selected Supplier")

        reason = assigned.get("Reason", "Fallback in split group")

        if selected_supplier != "No valid suppliers":
            pct_col = f"{selected_supplier} - Final % savings vs baseline"
            usd_col = f"{selected_supplier} - Final USD savings vs baseline"
            landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
            landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"

            fob_savings_pct = row.get(pct_col)
            fob_savings_usd = row.get(usd_col)
            landed_savings_pct = row.get(landed_pct_col)
            landed_savings_usd = row.get(landed_usd_col)

        other_supplier = row.get("2nd Lowest Bid Landed Supplier", "-")

    # --- Rule 1: Priority supplier logic (only if not split group) ---
    else:
        # Step 1: Retain incumbent if priority supplier
        if incumbent in priority_suppliers:
            selected_supplier = incumbent
            reason = f"Priority incumbent ({incumbent}) retained"
            pct_col = f"{incumbent} - Final % savings vs baseline"
            usd_col = f"{incumbent} - Final USD savings vs baseline"
            fob_savings_pct = row.get(pct_col)
            fob_savings_usd = row.get(usd_col)
            landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
            landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
            landed_savings_pct = row.get(landed_pct_col)
            landed_savings_usd = row.get(landed_usd_col)

        else:
            # Step 2: If any priority supplier has savings
            for supplier in priority_suppliers:
                pct_col = f"{supplier} - Final % savings vs baseline"
                if row[pct_col] == '-': continue
                if pct_col in row and pd.notna(row[pct_col]) and float(row[pct_col]) > 0:
                    selected_supplier = supplier
                    reason = f"Awarded to priority supplier ({supplier}) due to savings"
                    fob_savings_pct = row.get(pct_col)
                    usd_col = f"{supplier} - Final USD savings vs baseline"
                    fob_savings_usd = row.get(usd_col)
                    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
                    landed_savings_pct = row.get(landed_pct_col)
                    landed_savings_usd = row.get(landed_usd_col)
                    break

            # Step 3: If any priority supplier is lowest bidder
            if selected_supplier is None and min_supplier in priority_suppliers:
                selected_supplier = min_supplier
                reason = f"Lowest bid by priority supplier ({min_supplier})"
                pct_col = f"{min_supplier} - Final % savings vs baseline"
                usd_col = f"{min_supplier} - Final USD savings vs baseline"
                fob_savings_pct = row.get(pct_col)
                fob_savings_usd = row.get(usd_col)
                landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
                landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
                landed_savings_pct = row.get(landed_pct_col)
                landed_savings_usd = row.get(landed_usd_col)

            else:
                # Step 4: If no priority supplier, assign lowest bidder
                if pd.notna(min_supplier):
                    selected_supplier = min_supplier
                    reason = f"Assigned to lowest bidder ({min_supplier})"
                    pct_col = f"{min_supplier} - Final % savings vs baseline"
                    usd_col = f"{min_supplier} - Final USD savings vs baseline"
                    fob_savings_pct = row.get(pct_col)
                    fob_savings_usd = row.get(usd_col)
                    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
                    landed_savings_pct = row.get(landed_pct_col)
                    landed_savings_usd = row.get(landed_usd_col)

    
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
    def get_ref_value(col):
        return ref_row[col].values[0] if not ref_row.empty else "-"
    
    try:
        fob_savings_usd = float(fob_savings_usd) if pd.notna(fob_savings_usd) else 0
    except (ValueError, TypeError):
        fob_savings_usd = 0
    
    total_fob_savings_usd += fob_savings_usd 
        
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
            net_new_supplier_list.add(selected_supplier)
            net_new_supplier_count += 1
            total_landed_cost_completely_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)


    if selected_supplier != "-": unique_suppliers.add(selected_supplier) 

    unique_suppliers.add(selected_supplier)

    # --- Build output row ---
    output_row = {
        "ROW ID #": row_id,
        "Division": row.get("Division"),
        "Part #": part_no,
        "Item Description": row.get("Item Description"),
        "Product Group": row.get("Product Group"),
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "FOB Savings %": fob_savings_pct,
        "FOB Savings USD": fob_savings_usd,
        "Landed Cost Savings %": landed_savings_pct,
        "Landed Cost Savings USD": landed_savings_usd,
        "Reason": reason,
        "Landed Extended Cost USD": row[f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"] * row.get(volume_col, 0),
        "Is Totally New Supplier": "Yes" if (selected_supplier not in incumbent_suppliers and selected_supplier != '-') else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
        "Backup Supplier": other_supplier,
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
output_file = "Scenario_2_Output.xlsx"

# Write to Excel
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario - 2"
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

