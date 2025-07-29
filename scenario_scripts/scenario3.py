'''
Scenario 3:
- Grab the lowest bids for all parts from **new suppliers** (not the incumbent) until cumulative extended cost reaches **20% of 40.9M** (≈ $8.18M).
- Extended cost = Annual Volume (per UOM) × Supplier's R2 - Total landed cost per UOM (USD).
- A supplier is “new” if they are not equal to the Normalized incumbent supplier.
- After 20% is reached, remaining parts are awarded to the **incumbent** supplier.
- If the new supplier has not bid, fall back to incumbent. If incumbent also didn’t bid, fall back to “Final Minimum Bid Landed Supplier”.


10 % new new, 10 % new incumbent, 80 % incumbent.


- check if the supplier is overcommitted, by checking the how much a supplier has been awarded by calculating the extended cost of the supplier.


'''
import pandas as pd
from tqdm import tqdm
import time

# --- Start timer ---
start_time = time.time()

# --- Constants ---
PERCENT_NEW = 0.20  
TOTAL_COST = 40987650.3382
THRESHOLD_COST = TOTAL_COST * PERCENT_NEW

incumbent_col = "Normalized incumbent supplier"
valid_supplier_col = "Valid Supplier"
volume_col = "Annual Volume (per UOM)"

# --- Load files ---
input_path = "new/Bidsheet Master Consolidate Landed.csv"
output_reference_file_path = "new/outout-reference.csv"

print("Reading:", input_path)
df = pd.read_csv(input_path)
output_reference_df = pd.read_csv(output_reference_file_path)
print(f"✅ Loaded {len(df)} rows\n")

# --- Identify R2 landed cost columns ---
r2_fob_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]

# --- Prepare list for switching decisions ---
decision_rows = []

incumbent_suppliers = df[incumbent_col].unique()
suppliers = [col.split(" - R2")[0] for col in r2_fob_cols]
new_supplier_spent = 0
totally_new_suppliers = len(set(suppliers) - set(incumbent_suppliers))


print("Evaluating lowest new supplier bids with extended cost...\n")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Preparing"):

    valid_supplier_count = row.get(valid_supplier_col, 0)
    if valid_supplier_count == 0:
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": "-",
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "No valid suppliers"
        })
        continue

    min_bidder  = row.get("Final Minimum Bid Landed Supplier")
    if min_bidder != row.get(incumbent_col):
        extended_cost = row.get(f"{min_bidder} - R2 - Total landed cost per UOM (USD)", 0) * pd.to_numeric(row.get(volume_col), errors="coerce")
        if new_supplier_spent + extended_cost<= THRESHOLD_COST:
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": min_bidder,
                "extended_cost": extended_cost,
                "incumbent": row.get(incumbent_col),
                "reason": f"Switched to new supplier within {PERCENT_NEW*100}% threshold"
            })
            new_supplier_spent += extended_cost
        else:
            incumbent = row.get(incumbent_col)
            if incumbent in suppliers:
                decision_rows.append({
                    "index": idx,
                    "row": row,
                    "new_supplier": incumbent,
                    "extended_cost": 0,
                    "incumbent": incumbent,
                    "reason": f"{PERCENT_NEW*100}% Threshold exceeded, retaining incumbent"
                })
            else:
                supplier = row.get("Final Minimum Bid Landed Supplier", "-")
                decision_rows.append({
                    "index": idx,
                    "row": row,
                    "new_supplier": supplier,
                    "extended_cost": 0,
                    "incumbent": incumbent,
                    "reason": "Incumbent Supplier did not bid, Final Minimum Bid Landed Supplier used"
                })  

    else:
        incumbent = row.get(incumbent_col)
        if incumbent in suppliers:
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": incumbent,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": "Retaining incumbent supplier"
            })
        else:
            supplier = row.get("Final Minimum Bid Landed Supplier", "-")
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": supplier,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": "Incumbent Supplier did not bid, Final Minimum Bid Landed Supplier used"
            })

# --- Final output ---

output_data = []
total_fob_savings_usd = 0
total_landed_savings_usd = 0
incumbent_retained = 0
new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0

print("\nBuilding final output rows...\n")
for decision in tqdm(decision_rows, total=len(decision_rows), desc="Finalizing"):
    row = decision["row"]
    idx = decision["index"]
    selected_supplier = decision["new_supplier"]
    reason = decision["reason"]
    incumbent = decision["incumbent"]

    # Get savings columns
    pct_col = f"{selected_supplier} - Final % savings vs baseline"
    usd_col = f"{selected_supplier} - Final USD savings vs baseline"
    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"

    # Reference supplier metadata
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
    def get_ref_value(col):
        return ref_row[col].values[0] if not ref_row.empty else "-"
    
    try:
        fob_savings_usd = float(row.get(usd_col)) if pd.notna(row.get(usd_col)) else 0
    except (ValueError, TypeError):
        fob_savings_usd = 0
    
    total_fob_savings_usd += fob_savings_usd 
        
    try:
        lcs = float(row.get(landed_usd_col)) if pd.notna(landed_usd_col) else 0
    except (ValueError, TypeError): 
        lcs = 0

    total_landed_savings_usd  += lcs

    if selected_supplier == incumbent:
        incumbent_retained += 1
        total_landed_cost_incumbent += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
    else:
        new_supplier_count += 1

        if selected_supplier in incumbent_suppliers:
            total_landed_cost_new_suppliers += row.get(volume_col, 0) * row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        else:
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
        "Is Totally New Supplier": "Yes" if selected_supplier not in incumbent_suppliers else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
        "Standard leadtime - days PO-shipment POL": get_ref_value("Standard leadtime - days PO-shipment POL"),
        "Retail Packaging": get_ref_value("Retail Packaging"),
        "Payment term - days and discounts": get_ref_value("Payment term - days and discounts"),
        "New product introduction": get_ref_value("New product introduction"),
        "Long term commitment rebate": get_ref_value("Long term commitment rebate"),
        "Uncompetitive supplier behavior": get_ref_value("Uncompetitive supplier behavior"),
    }
    output_data.append(output_row)

summary_data = [
    ["Total FOB Savings USD", total_fob_savings_usd],
    ["Total Landed Cost Savings USD", total_landed_savings_usd],
    ["Total parts where Incumbent Suppliers Retained", incumbent_retained],
    ["Total Landed Cost where Incumbent Suppliers Retained", total_landed_cost_incumbent],
    ["Total parts where bid is awarded to New Suppliers", new_supplier_count],
    ["Total Landed Cost where bid is awarded to New Suppliers", total_landed_cost_new_suppliers],
    ["Total Landed Cost where bid is awarded to Completely New Suppliers", total_landed_cost_completely_new_suppliers],
    ["Total Unique Suppliers", len(unique_suppliers)],
    ["Totally New Suppliers", totally_new_suppliers],
]
# Output file
output_file = "Scenario_3_20_Output.xlsx"

# --- Write to Excel ---
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Write summary
    for row_idx, row in enumerate(summary_data):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, col_idx, value)

    worksheet.write(9, 0, "Total Cost")
    worksheet.write_formula(9, 1, '=SUM(B2+B4+B6+B7)')

    # Write output table
    df_output = pd.DataFrame(output_data)
    df_output.to_excel(writer, sheet_name="Sheet1", startrow=len(summary_data) + 3, index=False)

# --- Timer ---
elapsed_time = time.time() - start_time
print(f"\n✅ Done. Output written to '{output_file}'")
print(f"⏱ Time taken: {elapsed_time:.2f} seconds")