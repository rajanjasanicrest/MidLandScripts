
'''
Scenario 3:
- Grab the lowest bids for all parts from **new suppliers** (not the incumbent) until cumulative extended cost reaches **20% of 40.9M** (≈ $8.18M).
- Extended cost = Annual Volume (per UOM) × Supplier's R2 - Total landed cost per UOM (USD).
- A supplier is “new” if they are not equal to the Normalized incumbent supplier.
- After 20% is reached, remaining parts are awarded to the *incumbent* supplier.
- If the new supplier has not bid, fall back to incumbent. If incumbent also didn’t bid, fall back to “Final Minimum Bid Landed Supplier”.
'''
import pandas as pd
from tqdm import tqdm
import time

# --- Start timer ---
start_time = time.time()

# --- Constants ---
PERCENT_NEW = 0.40
TOTAL_COST = 49783721.6480354
THRESHOLD_COST = TOTAL_COST * PERCENT_NEW

incumbent_col = "Normalized incumbent supplier"
valid_supplier_col = "Valid Supplier"
volume_col = "Annual Volume (per UOM)"

# --- Load files ---
input_path = "new/Bidsheet Master Consolidate Landed3.csv"
output_reference_file_path = "new/outout-reference.csv"

print("Reading:", input_path)
df = pd.read_csv(input_path)
output_reference_df = pd.read_csv(output_reference_file_path)
print(f"Loaded {len(df)} rows\n")

# --- Load discount data from CSV ---
discount_df = pd.read_csv("new/discount.csv")
print(f"Loaded discount data: {len(discount_df)} rows")

def parse_range(range_str):
    """Parse range string like '100-150' or '1,000-1,250' to get min and max values"""
    if pd.isna(range_str) or range_str == "":
        return None, None
    
    # Remove commas and split by dash
    range_clean = str(range_str).replace(",", "")
    if "-" in range_clean:
        try:
            parts = range_clean.split("-")
            min_val = float(parts[0])
            max_val = float(parts[1])
            return min_val, max_val
        except:
            return None, None
    return None, None

def get_discount_for_supplier_amount(supplier_name, annual_fob_cost_thousands):
    """Get discount percentage for supplier based on annual FOB cost in thousands USD"""
    supplier_discounts = discount_df[discount_df['Supplier Name'] == supplier_name]
    
    if supplier_discounts.empty:
        return 0.0
    
    # Find the appropriate tier
    for _, row in supplier_discounts.iterrows():
        min_val, max_val = parse_range(row['Annual Revenue  Requirement in 1,000 USD'])
        if min_val is not None and max_val is not None:
            if min_val <= annual_fob_cost_thousands <= max_val:
                discount_pct = row['% Discount off EXW Price']
                if pd.isna(discount_pct) or discount_pct == "":
                    return 0.0
                try:
                    return float(discount_pct)
                except:
                    return 0.0
    
    return 0.0


# --- Identify R2 landed cost columns ---
r2_fob_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]

# --- Prepare suppliers ---
incumbent_suppliers = df[incumbent_col].unique()
suppliers = [col.split(" - R2")[0] for col in r2_fob_cols]

# --- PART ASSIGNMENT LOGIC (HEAVILY COMMENTED) ---
# We process all rows and classify them into:
#   1. No valid suppliers: Not awarded.
#   2. Incumbent did not bid, but minimum bid exists: Assign to min bid (contributes to 20% threshold).
#   3. Incumbent bid:
#       a. Incumbent is the minimum: Retain incumbent (does NOT contribute to 20% threshold).
#       b. Incumbent is NOT the minimum: Assign to min bid (contributes to 20% threshold).

no_valid_supplier_parts = []
must_assign_min_bid_parts = []  # Forced to min bid, always contributes to threshold
incumbent_retained_parts = []
candidate_new_supplier_parts = []  # Eligible for threshold assignment
net_new_supplier_list = set()
total_cost_not_awarded = 0

for idx, row in df.iterrows():
    incumbent = row.get(incumbent_col)
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
    valid_supplier_count = row.get(valid_supplier_col, 0)

    # 1. No valid suppliers
    if valid_supplier_count == 0 or pd.isna(min_supplier):
        no_valid_supplier_parts.append({
            "index": idx,
            "row": row,
            "reason": "No valid suppliers"
        })
        total_cost_not_awarded += row.get('Landed Extended Cost USD', 0)
        continue

    # 2. Incumbent did not bid, but minimum bid exists
    if incumbent not in suppliers and pd.notna(min_supplier):
        landed_cost = row.get(f"{min_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        volume = pd.to_numeric(row.get(volume_col), errors='coerce')
        extended_cost = landed_cost * volume
        savings_usd = row.get(f"{min_supplier} - Final Landed USD savings vs baseline", 0)
        try:
            savings_usd = float(savings_usd)
        except:
            savings_usd = 0
        must_assign_min_bid_parts.append({
            "index": idx,
            "row": row,
            "extended_cost": extended_cost,
            "savings_usd": savings_usd,
            "min_supplier": min_supplier,
            "incumbent": incumbent,
            "reason": "Incumbent did not bid, using Final Minimum Bid Landed Supplier"
        })
        continue

    # 3. Incumbent bid
    incumbent_bid_val = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
    if incumbent_bid_val > 0:
        # a. Incumbent is the minimum
        if min_supplier == incumbent:
            incumbent_retained_parts.append({
                "index": idx,
                "row": row,
                "incumbent": incumbent,
                "reason": "Incumbent retained (lowest bid)"
            })
        # b. Incumbent is NOT the minimum
        else:
            landed_cost = row.get(f"{min_supplier} - R2 - Total landed cost per UOM (USD)", 0)
            volume = pd.to_numeric(row.get(volume_col), errors='coerce')
            extended_cost = landed_cost * volume
            savings_usd = row.get(f"{min_supplier} - Final Landed USD savings vs baseline", 0)
            try:
                savings_usd = float(savings_usd)
            except:
                savings_usd = 0
            candidate_new_supplier_parts.append({
                "index": idx,
                "row": row,
                "extended_cost": extended_cost,
                "savings_usd": savings_usd,
                "min_supplier": min_supplier,
                "incumbent": incumbent,
                "reason": "Incumbent bid, but not lowest; eligible for new supplier assignment"
            })
    else:
        # Incumbent did not bid, but minimum bid exists (should already be handled above)
        pass

# --- Sort candidate new supplier parts by savings descending ---
candidate_new_supplier_parts.sort(key=lambda x: x["savings_usd"], reverse=True)

# --- Assign must-assign-min-bid parts first (these are forced, contribute to threshold) ---
decision_rows = []
new_supplier_spent = 0
selected_new_rows = set()

for part in must_assign_min_bid_parts:
    if new_supplier_spent + part["extended_cost"] <= THRESHOLD_COST:
        decision_rows.append({
            "index": part["index"],
            "row": part["row"],
            "new_supplier": part["min_supplier"],
            "extended_cost": part["extended_cost"],
            "incumbent": part["incumbent"],
            "reason": part["reason"]
        })
        new_supplier_spent += part["extended_cost"]
        selected_new_rows.add(part["index"])
    else:
        # If threshold exceeded, assign to incumbent (if possible)
        decision_rows.append({
            "index": part["index"],
            "row": part["row"],
            "new_supplier": part["incumbent"],
            "extended_cost": 0,
            "incumbent": part["incumbent"],
            "reason": "Threshold exceeded, retaining incumbent"
        })
        selected_new_rows.add(part["index"])

# --- Assign candidate new supplier parts until threshold hit ---
for part in candidate_new_supplier_parts:
    if part["index"] in selected_new_rows:
        continue
    if new_supplier_spent + part["extended_cost"] <= THRESHOLD_COST:
        decision_rows.append({
            "index": part["index"],
            "row": part["row"],
            "new_supplier": part["min_supplier"],
            "extended_cost": part["extended_cost"],
            "incumbent": part["incumbent"],
            "reason": part["reason"] + f" (within {PERCENT_NEW*100}% threshold)"
        })
        new_supplier_spent += part["extended_cost"]
        selected_new_rows.add(part["index"])
    else:
        # Threshold exceeded, retain incumbent
        decision_rows.append({
            "index": part["index"],
            "row": part["row"],
            "new_supplier": part["incumbent"],
            "extended_cost": 0,
            "incumbent": part["incumbent"],
            "reason": "Threshold exceeded, retaining incumbent"
        })
        selected_new_rows.add(part["index"])

# --- Assign all incumbent retained parts ---
for part in incumbent_retained_parts:
    if part["index"] in selected_new_rows:
        continue
    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": part["incumbent"],
        "extended_cost": 0,
        "incumbent": part["incumbent"],
        "reason": part["reason"]
    })
    selected_new_rows.add(part["index"])

# --- Assign all no valid supplier parts ---
for part in no_valid_supplier_parts:
    if part["index"] in selected_new_rows:
        continue
    decision_rows.append({
        "index": part["index"],
        "row": part["row"],
        "new_supplier": "-",
        "extended_cost": 0,
        "incumbent": part["row"].get(incumbent_col),
        "reason": part["reason"]
    })
    selected_new_rows.add(part["index"])

# --- Step 4: Assign rest (fallback to incumbent or final bid supplier) ---
for idx, row in df.iterrows():
    if idx in selected_new_rows:
        continue

    incumbent = row.get(incumbent_col)
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
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
    elif incumbent in suppliers:
        incumbent_bid_val = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
        if incumbent_bid_val > 0:
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": incumbent,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": f"Outside {PERCENT_NEW*100}% threshold, retaining incumbent"
            })
        else:
            decision_rows.append({
                "index": idx,
                "row": row,
                "new_supplier": min_supplier,
                "extended_cost": 0,
                "incumbent": incumbent,
                "reason": f"Forced to Lowest Bidder"
            })
    elif pd.notna(min_supplier):
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": min_supplier,
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "Incumbent did not bid, using Final Minimum Bid Landed Supplier"
        })
    else:
        decision_rows.append({
            "index": idx,
            "row": row,
            "new_supplier": "-",
            "extended_cost": 0,
            "incumbent": incumbent,
            "reason": "No valid bids"
        })

# --- Calculate annual revenue discounts per supplier ---
def calculate_annual_revenue_discounts():
    """Calculate annual revenue discounts for each supplier based on their total FOB cost"""
    supplier_fob_totals = {}
    supplier_parts = {}
    
    # First pass: calculate total FOB cost per supplier
    for decision in decision_rows:
        row = decision["row"]
        selected_supplier = decision["new_supplier"]
        
        if selected_supplier == "-":
            continue
            
        # Get FOB cost column
        fob_col = f"{selected_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
        fob_cost_per_uom = row.get(fob_col, 0)
        
        if pd.isna(fob_cost_per_uom):
            fob_cost_per_uom = 0
        
        volume = pd.to_numeric(row.get(volume_col), errors='coerce')
        if pd.isna(volume):
            volume = 0
            
        # Calculate annual FOB cost (multiply by 0.8 to convert 15 months to 12 months)
        annual_fob_cost = fob_cost_per_uom * volume * 0.8
        
        if selected_supplier not in supplier_fob_totals:
            supplier_fob_totals[selected_supplier] = 0
            supplier_parts[selected_supplier] = []
            
        supplier_fob_totals[selected_supplier] += annual_fob_cost
        supplier_parts[selected_supplier].append({
            'index': decision["index"],
            'annual_fob_cost': annual_fob_cost
        })
    
    # Second pass: calculate discount per supplier and distribute
    supplier_discount_amounts = {}
    
    for supplier, total_annual_fob in supplier_fob_totals.items():
        # Convert to thousands for discount lookup
        annual_fob_thousands = total_annual_fob / 1000
        
        # Get discount rate from CSV data
        discount_rate = get_discount_for_supplier_amount(supplier, annual_fob_thousands)
        total_discount = total_annual_fob * discount_rate
        
        # Distribute discount proportionally across parts
        parts = supplier_parts[supplier]
        supplier_discount_amounts[supplier] = {}
        
        if total_annual_fob > 0:
            for part in parts:
                part_proportion = part['annual_fob_cost'] / total_annual_fob
                part_discount = total_discount * part_proportion
                supplier_discount_amounts[supplier][part['index']] = part_discount
        else:
            for part in parts:
                supplier_discount_amounts[supplier][part['index']] = 0
    
    return supplier_discount_amounts

# Calculate annual revenue discounts
annual_revenue_discounts = calculate_annual_revenue_discounts()

# --- Final output processing ---
output_data = []
total_fob_savings_usd = 0
total_landed_savings_usd = 0
total_annual_revenue_discount = 0
incumbent_retained = 0
new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0

net_new_supplier_count = 0  
parts_where_no_bids = 0

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
    
    # Get annual revenue discount for this part
    annual_discount = 0
    if selected_supplier != "-" and selected_supplier in annual_revenue_discounts:
        annual_discount = annual_revenue_discounts[selected_supplier].get(idx, 0)
    
    total_annual_revenue_discount += annual_discount
    
    if selected_supplier == "-":
        # parts_where_no_bids += 1
        output_row = {
            "ROW ID #": row.get("ROW ID #"),
            "Division": row.get("Division"),
            "Part #": row.get("Part #"),
            "Item Description": row.get("Item Description"),
            "Product Group": row.get("Product Group"),
            "Part Family": row.get("Part Family"),
            "Incumbent Supplier": incumbent,
            "Selected Supplier": incumbent,
            "Annual Volume (per UOM)": row.get("Annual Volume (per UOM)"),
            "FOB Savings %": "-",
            "FOB Savings USD": "-",
            "Landed Cost Savings %": "-",
            "Landed Cost Savings USD": "-",
            "Annual revenue discount USD": 0,
            "Reason": 'No valid bids in this round, so forced to incumbent supplier.',
            "Landed Extended Cost USD": row.get('Landed Extended Cost USD'),
            "Is Totally New Supplier": "No",
            "Part Switched": "No",
            "Standard leadtime - days PO-shipment POL": get_ref_value("Standard leadtime - days PO-shipment POL"),
            "Retail Packaging": get_ref_value("Retail Packaging"),
            "Payment term - days and discounts": get_ref_value("Payment term - days and discounts"),
            "New product introduction": get_ref_value("New product introduction"),
            "Long term commitment rebate": get_ref_value("Long term commitment rebate"),
            "Uncompetitive supplier behavior": get_ref_value("Uncompetitive supplier behavior"),
            "valid_supplier_count": row.get('Valid Supplier')
        }
        output_data.append(output_row)
        continue

    # Note: Supplier metrics will be recalculated after rationalization
    # to ensure accuracy after any supplier reassignments

    output_row = {
        "Annual revenue discount USD": annual_discount,
        "ROW ID #": row.get("ROW ID #"),
        "Division": row.get("Division"),
        "Part #": row.get("Part #"),
        "Item Description": row.get("Item Description"),
        "Product Group": row.get("Product Group"),
        "Part Family": row.get("Part Family"),
        "Incumbent Supplier": row.get("Normalized incumbent supplier"),
        "Selected Supplier": selected_supplier,
        "Annual Volume (per UOM)": row.get('Annual Volume (per UOM)'),
        "FOB Savings %": row.get(pct_col, "-"),
        "FOB Savings USD": row.get(usd_col, "-"),
        "Landed Cost Savings %": row.get(landed_pct_col, "-"),
        "Landed Cost Savings USD": row.get(landed_usd_col, "-"),
        "Reason": reason,
        "Landed Extended Cost USD": row[f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"] * row.get(volume_col, 0),
        "Is Totally New Supplier": "Yes" if selected_supplier not in incumbent_suppliers else "No",
        "Part Switched": "Yes" if selected_supplier != incumbent else "No",
        "Standard leadtime - days PO-shipment POL": get_ref_value("Standard leadtime - days PO-shipment POL"),
        "Retail Packaging": get_ref_value("Retail Packaging"),
        "Payment term - days and discounts": get_ref_value("Payment term - days and discounts"),
        "New product introduction": get_ref_value("New product introduction"),
        "Long term commitment rebate": get_ref_value("Long term commitment rebate"),
        "Uncompetitive supplier behavior": get_ref_value("Uncompetitive supplier behavior"),
        "valid_supplier_count": row.get('Valid Supplier')
    }
    output_data.append(output_row)

# --- BINZHOU ZELI1 REMOVAL LOGIC (BEFORE RATIONALIZATION) ---
print("\nApplying Binzhou Zeli removal logic for specific parts...")

binzhou_zeli_supplier = "Binzhou Zeli"
binzhou_reassignments = 0

# Define specific part numbers where Binzhou Zeli should be avoided
# binzhou_zeli_exclusion_parts = [
#     "DDSL-2020-A1", "DDSL-3030-A1",
#     "DDSL-4040-A1",
# ]
binzhou_zeli_exclusion_parts = [
    "CDCSL-200-A1", "CDCSL-300-A1", "CDCSL-400-A1", "CDCSL-600-A1",
    "CGBSL-300-A1", "CGBSL-400-A1", "CGCSL-200CR-A1", "CGCSL-300CR-A1",
    "CGCSL-400CR-A1", "CGCSL-600CR-A1", "CGDSL-200-A1", "CGDSL-300-A1",
    "CGDSL-400-A1", "CGDSL-600-A1", "DASL-3020-A1", "DASL-3040-A1",
    "DASL-4030-A1", "DASL-6040-A1", "DDSL-2020-A1", "DDSL-3030-A1",
    "DDSL-4040-A1", "DASL-2030-A1", "CGBSL-200-A1"
]

print(f"Binzhou Zeli exclusion applies to {len(binzhou_zeli_exclusion_parts)} specific part numbers")

def find_best_alternative_to_binzhou(row, all_suppliers):
    """Find the best alternative supplier excluding Binzhou Zeli"""
    incumbent = row.get("Incumbent Supplier", "")
    valid_supplier_count = row.get("Valid Supplier", 0)
    
    # If only one valid supplier and it's Binzhou Zeli, we have no choice
    if valid_supplier_count == 1:
        # Check if Binzhou Zeli is the only bidder
        binzhou_landed_cost = row.get(f"{binzhou_zeli_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.notna(binzhou_landed_cost) and binzhou_landed_cost > 0:
            # Count other valid bidders
            other_valid_bidders = 0
            for supplier in all_suppliers:
                if supplier != binzhou_zeli_supplier:
                    landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
                    if pd.notna(landed_cost) and landed_cost > 0:
                        other_valid_bidders += 1
            
            if other_valid_bidders == 0:
                return binzhou_zeli_supplier, "Only valid supplier available"
    
    # First preference: incumbent (if not Binzhou Zeli and has valid bid)
    if incumbent != binzhou_zeli_supplier and incumbent in all_suppliers:
        incumbent_landed_cost = row.get(f"{incumbent} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.notna(incumbent_landed_cost) and incumbent_landed_cost > 0:
            return incumbent, "Reassigned to incumbent (avoiding Binzhou Zeli)"
    
    # Second preference: find lowest bidder excluding Binzhou Zeli
    best_supplier = None
    best_cost = float('inf')
    
    for supplier in all_suppliers:
        if supplier == binzhou_zeli_supplier:
            continue
            
        landed_cost = row.get(f"{supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.notna(landed_cost) and landed_cost > 0 and landed_cost < best_cost:
            best_cost = landed_cost
            best_supplier = supplier
    
    if best_supplier:
        return best_supplier, f"Reassigned to lowest bidder excluding Binzhou Zeli (${best_cost:.2f})"
    
    # Fallback: if no other valid bidders, keep Binzhou Zeli
    return binzhou_zeli_supplier, "No alternative suppliers available"

# Get all R2 landed cost columns for finding alternatives
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]

# Create a lookup dictionary for faster DataFrame access (highly optimized)
print("Creating DataFrame lookup for performance optimization...")
df_lookup = {}
for idx, row in df.iterrows():
    row_id = row.get("ROW ID #")
    if row_id is not None:
        df_lookup[row_id] = row
print(f"DataFrame lookup created with {len(df_lookup)} entries")

# Apply Binzhou Zeli removal logic to output_data (only for specific parts)
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    part_number = row.get("Part #", "")
    
    # Only apply Binzhou Zeli removal logic to specific part numbers
    if current_supplier == binzhou_zeli_supplier and part_number in binzhou_zeli_exclusion_parts:
        # Find corresponding row in original dataframe using lookup
        row_id = row.get("ROW ID #")
        df_row = df_lookup.get(row_id)
        
        if df_row is not None:
            new_supplier, reason = find_best_alternative_to_binzhou(df_row, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"Binzhou Zeli avoided: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row.get(landed_pct_col, "-")
                    output_data[i]["Landed Cost Savings USD"] = df_row.get(landed_usd_col, "-")
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    landed_cost = df_row.get(f"{new_supplier} - R2 - Total landed cost per UOM (USD)", 0)
                    output_data[i]["Landed Extended Cost USD"] = landed_cost * volume
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                    
                    output_data[i]["Standard leadtime - days PO-shipment POL"] = get_ref_value("Standard leadtime - days PO-shipment POL")
                    output_data[i]["Retail Packaging"] = get_ref_value("Retail Packaging")
                    output_data[i]["Payment term - days and discounts"] = get_ref_value("Payment term - days and discounts")
                    output_data[i]["New product introduction"] = get_ref_value("New product introduction")
                    output_data[i]["Long term commitment rebate"] = get_ref_value("Long term commitment rebate")
                    output_data[i]["Uncompetitive supplier behavior"] = get_ref_value("Uncompetitive supplier behavior")
                
                binzhou_reassignments += 1

print(f"Binzhou Zeli removal complete: {binzhou_reassignments} parts reassigned")

# Update decision_rows with Binzhou Zeli changes (optimized)
output_row_lookup = {row.get("ROW ID #"): row for row in output_data}
for i, decision in enumerate(decision_rows):
    row_id = decision["row"].get("ROW ID #")
    if row_id in output_row_lookup:
        output_row = output_row_lookup[row_id]
        decision_rows[i]["new_supplier"] = output_row["Selected Supplier"]
        if "Binzhou Zeli avoided" in output_row["Reason"]:
            decision_rows[i]["reason"] = output_row["Reason"]

# Recalculate annual revenue discounts after Binzhou Zeli removal
print("Recalculating annual revenue discounts after Binzhou Zeli removal...")
annual_revenue_discounts = calculate_annual_revenue_discounts()

# Update annual revenue discount in output_data after Binzhou Zeli removal (optimized)
total_annual_revenue_discount = 0
decision_row_lookup = {decision["row"].get("ROW ID #"): i for i, decision in enumerate(decision_rows)}

for i, row in enumerate(output_data):
    selected_supplier = row["Selected Supplier"]
    row_id = row.get("ROW ID #")
    
    # Use lookup instead of nested loop
    decision_idx = decision_row_lookup.get(row_id)
    
    annual_discount = 0
    if selected_supplier != "-" and selected_supplier in annual_revenue_discounts and decision_idx is not None:
        annual_discount = annual_revenue_discounts[selected_supplier].get(decision_idx, 0)
    
    output_data[i]["Annual revenue discount USD"] = annual_discount
    total_annual_revenue_discount += annual_discount

print(f"Annual revenue discounts recalculated after Binzhou Zeli removal. Total: ${total_annual_revenue_discount:,.2f}")

# --- TAIL SUPPLIER RATIONALIZATION LOGIC (AFTER BINZHOU ZELI REMOVAL) ---
print("\nApplying tail supplier rationalization logic...\n")

# Calculate total awarded amount per supplier
supplier_awarded_amounts = {}
for row in output_data:
    supplier = row["Selected Supplier"]
    valid_supplier_count = row["valid_supplier_count"]
    # print(valid_supplier_count)
    if valid_supplier_count != 0:
        awarded_amount = row.get("Landed Extended Cost USD", 0)
        if supplier not in supplier_awarded_amounts:
            supplier_awarded_amounts[supplier] = 0
        supplier_awarded_amounts[supplier] += awarded_amount

# Dynamically identify tail suppliers (those with <$100k total awarded)
tail_suppliers_to_rationalize = ['Giraffe Stainless', 'Union Metal Products', 'WEFLO', 'Kaixuan Stainless Steel', 'Tianjin Outshine', 'Sichuan Y&J', 'Guangzhou Hopetrol', 'Swati Enterprise']
# Identify suppliers with ≥$100k awards (eligible to receive rationalized parts)
large_suppliers = {supplier: amount for supplier, amount in supplier_awarded_amounts.items() 
                  if amount >= 100000}

print(f"Large suppliers (≥$100k): {len(large_suppliers)}")
for supplier, amount in sorted(large_suppliers.items(), key=lambda x: x[1], reverse=True):
    print(f"  - {supplier}: ${amount:,.2f}")

print(f"\nTail suppliers (<$100k) to rationalize: {len(tail_suppliers_to_rationalize)}")
for supplier, amount in sorted([(s, supplier_awarded_amounts[s]) for s in tail_suppliers_to_rationalize], 
                              key=lambda x: x[1], reverse=True):
    print(f"  - {supplier}: ${amount:,.2f}")

# Get all R2 landed cost columns for finding next best bidders
r2_landed_cols = [col for col in df.columns if col.endswith("R2 - Total landed cost per UOM (USD)")]
all_suppliers = [col.split(" - R2")[0] for col in r2_landed_cols]

def find_next_best_large_supplier(row, current_supplier, large_suppliers, all_suppliers):
    """Find the next best bidder among large suppliers for a given part"""

    incumbent = row.get("Normalized incumbent supplier", "")
    valid_supplier = row.get('Valid Supplier')
    if valid_supplier == 1:
        if incumbent == current_supplier:
            return current_supplier, "Has to stay with it."
        else:
            return incumbent, 'Forced to incumbent because no other bid on it.'
    else:

        incumb_bid_col = f'{incumbent} - R2 - Total landed cost per UOM (USD)'
        incum_bid = row.get(incumb_bid_col, 0)
        
        part_bids = []
        # First check if incumbent is a large supplier
        if incumbent in large_suppliers:
            if pd.notna(incum_bid) and incum_bid > 0:
                part_bids.append((incumbent, incum_bid))
                return incumbent, "Rationalized to other bidder than than bidder based on logic"
        
        # Get all bids for this part and sort by landed cost
        for supplier in all_suppliers:
            if supplier == current_supplier:
                continue  # Skip current tail supplier
                
            landed_cost_col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
            landed_cost = row.get(landed_cost_col, 0)
            
            if pd.notna(landed_cost) and landed_cost > 0:
                part_bids.append((supplier, landed_cost))
        
        # Sort by landed cost (ascending - lowest cost first)
        part_bids.sort(key=lambda x: x[1])
        
        # Find first large supplier in sorted list
        for supplier, cost in part_bids:
            if supplier in large_suppliers:
                return supplier, f"Rationalized to other bidder than than bidder based on logic"
        
        # If no large supplier found, return the incumbent anyway
        return incumbent, "incumbent anyway"

# Apply rationalization

rationalization_changes = 0
for i, row in enumerate(output_data):
    current_supplier = row["Selected Supplier"]
    
    if current_supplier in tail_suppliers_to_rationalize:
        # Find corresponding row in original dataframe
        row_id = row.get("ROW ID #")
        df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
        
        if df_row is not None and df_row['Valid Supplier'] >= 1:

            new_supplier, reason = find_next_best_large_supplier(df_row, current_supplier, large_suppliers, all_suppliers)
            
            if new_supplier != current_supplier:
                # Update the supplier assignment
                old_supplier = current_supplier
                output_data[i]["Selected Supplier"] = new_supplier
                output_data[i]["Reason"] = f"Rationalized from {old_supplier}: {reason}"
                
                # Update other relevant fields
                if new_supplier != "-":
                    # Update savings columns
                    pct_col = f"{new_supplier} - Final % savings vs baseline"
                    usd_col = f"{new_supplier} - Final USD savings vs baseline"
                    landed_pct_col = f"{new_supplier} - Final Landed % savings vs baseline"
                    landed_usd_col = f"{new_supplier} - Final Landed USD savings vs baseline"
                    
                    output_data[i]["FOB Savings %"] = df_row.get(pct_col, "-")
                    output_data[i]["FOB Savings USD"] = df_row.get(usd_col, "-")
                    output_data[i]["Landed Cost Savings %"] = df_row.get(landed_pct_col, "-")
                    output_data[i]["Landed Cost Savings USD"] = df_row.get(landed_usd_col, "-")
                    
                    # Update landed extended cost
                    volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
                    if pd.isna(volume):
                        volume = 0
                    landed_cost = df_row.get(f"{new_supplier} - R2 - Total landed cost per UOM (USD)", 0)
                    output_data[i]["Landed Extended Cost USD"] = landed_cost * volume
                    
                    # Update supplier classification
                    incumbent = df_row.get("Normalized incumbent supplier")
                    output_data[i]["Is Totally New Supplier"] = "Yes" if new_supplier not in incumbent_suppliers else "No"
                    output_data[i]["Part Switched"] = "Yes" if new_supplier != incumbent else "No"
                    
                    # Update reference data
                    ref_row = output_reference_df[output_reference_df['Reference'] == new_supplier]
                    def get_ref_value(col):
                        return ref_row[col].values[0] if not ref_row.empty else "-"
                    
                    output_data[i]["Standard leadtime - days PO-shipment POL"] = get_ref_value("Standard leadtime - days PO-shipment POL")
                    output_data[i]["Retail Packaging"] = get_ref_value("Retail Packaging")
                    output_data[i]["Payment term - days and discounts"] = get_ref_value("Payment term - days and discounts")
                    output_data[i]["New product introduction"] = get_ref_value("New product introduction")
                    output_data[i]["Long term commitment rebate"] = get_ref_value("Long term commitment rebate")
                    output_data[i]["Uncompetitive supplier behavior"] = get_ref_value("Uncompetitive supplier behavior")
                
                rationalization_changes += 1

print(f"Rationalization complete: {rationalization_changes} parts reassigned from tail suppliers")

# Recalculate annual revenue discounts after rationalization
print("Recalculating annual revenue discounts after rationalization...")

# Update decision_rows with rationalized assignments
for i, decision in enumerate(decision_rows):
    row_id = decision["row"].get("ROW ID #")
    # Find corresponding output row
    for output_row in output_data:
        if output_row.get("ROW ID #") == row_id:
            decision_rows[i]["new_supplier"] = output_row["Selected Supplier"]
            if "Rationalized from" in output_row["Reason"]:
                decision_rows[i]["reason"] = output_row["Reason"]
            break

# Recalculate annual revenue discounts with updated assignments
annual_revenue_discounts = calculate_annual_revenue_discounts()

# Update annual revenue discount in output_data
total_annual_revenue_discount = 0
for i, row in enumerate(output_data):
    selected_supplier = row["Selected Supplier"]
    row_id = row.get("ROW ID #")
    
    # Find the index in decision_rows
    decision_idx = None
    for j, decision in enumerate(decision_rows):
        if decision["row"].get("ROW ID #") == row_id:
            decision_idx = j
            break
    
    annual_discount = 0
    if selected_supplier != "-" and selected_supplier in annual_revenue_discounts and decision_idx is not None:
        annual_discount = annual_revenue_discounts[selected_supplier].get(decision_idx, 0)
    
    output_data[i]["Annual revenue discount USD"] = annual_discount
    total_annual_revenue_discount += annual_discount

print(f"Annual revenue discounts recalculated. Total: ${total_annual_revenue_discount:,.2f}")

# --- RECALCULATE ALL METRICS AFTER RATIONALIZATION ---
print("Recalculating all metrics after rationalization...")

# Reset all metrics
incumbent_retained = 0
new_supplier_count = 0
net_new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0
net_new_supplier_list = set()
total_landed_savings_usd = 0
total_fob_savings_usd = 0
# parts_where_no_bids = 0

# Recalculate all metrics based on final supplier assignments
for row in output_data:
    selected_supplier = row["Selected Supplier"]
    incumbent = row["Incumbent Supplier"]
    
    # Handle parts with no bids
    if selected_supplier == "-":
        parts_where_no_bids += 1
        # Add to total cost not awarded (use original landed extended cost from output_data)
        total_cost_not_awarded += row.get("Landed Extended Cost USD", 0)
        continue
    
    # Get volume and costs for calculations
    row_id = row.get("ROW ID #")
    df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
    
    if df_row is not None:
        volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
        if pd.isna(volume):
            volume = 0
        
        landed_cost = df_row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.isna(landed_cost):
            landed_cost = 0
        
        extended_cost = landed_cost * volume
        
        # Calculate savings
        fob_usd_col = f"{selected_supplier} - Final USD savings vs baseline"
        landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
        
        try:
            fob_savings = float(df_row.get(fob_usd_col, 0)) if pd.notna(df_row.get(fob_usd_col)) else 0
        except (ValueError, TypeError):
            fob_savings = 0
        
        try:
            landed_savings = float(df_row.get(landed_usd_col, 0)) if pd.notna(df_row.get(landed_usd_col)) else 0
        except (ValueError, TypeError):
            landed_savings = 0
        
        total_fob_savings_usd += fob_savings
        total_landed_savings_usd += landed_savings
        
        # Classify supplier and update metrics
        if selected_supplier == incumbent:
            incumbent_retained += 1
            total_incumbent_volume += volume
            total_landed_cost_incumbent += extended_cost
        else:
            if selected_supplier in incumbent_suppliers:
                new_supplier_count += 1
                total_new_supplier_volume += volume
                total_landed_cost_new_suppliers += extended_cost
            else:
                net_new_supplier_count += 1
                net_new_supplier_list.add(selected_supplier)
                total_net_new_supplier_volume += volume
                total_landed_cost_completely_new_suppliers += extended_cost
        
        unique_suppliers.add(selected_supplier)

print(f"All metrics recalculated after rationalization:")
print(f"  - Total landed savings USD: ${total_landed_savings_usd:,.2f}")
print(f"  - Total FOB savings USD: ${total_fob_savings_usd:,.2f}")
print(f"  - Total cost not awarded: ${total_cost_not_awarded:,.2f}")
print(f"  - Total landed cost incumbent: ${total_landed_cost_incumbent:,.2f}")
print(f"  - Total landed cost new suppliers: ${total_landed_cost_new_suppliers:,.2f}")
print(f"  - Total landed cost completely new suppliers: ${total_landed_cost_completely_new_suppliers:,.2f}")
print(f"  - Incumbent retained: {incumbent_retained}")
print(f"  - New suppliers: {new_supplier_count}")
print(f"  - Net new suppliers: {net_new_supplier_count}")
print(f"  - Parts with no bids: {parts_where_no_bids}")
print(f"  - Unique suppliers: {len(unique_suppliers)}")

# Final metrics recalculation after Binzhou Zeli removal
print("Final metrics recalculation after Binzhou Zeli removal...")

# Reset all metrics again
incumbent_retained = 0
new_supplier_count = 0
net_new_supplier_count = 0
unique_suppliers = set()
total_landed_cost_incumbent = 0
total_landed_cost_new_suppliers = 0
total_landed_cost_completely_new_suppliers = 0
total_incumbent_volume = 0
total_net_new_supplier_volume = 0
total_new_supplier_volume = 0
net_new_supplier_list = set()
total_landed_savings_usd = 0
total_fob_savings_usd = 0

# Final recalculation
for row in output_data:
    selected_supplier = row["Selected Supplier"]
    incumbent = row["Incumbent Supplier"]
    
    # Handle parts with no bids
    # if selected_supplier == "-":
    #     parts_where_no_bids += 1
    #     total_cost_not_awarded += row.get("Landed Extended Cost USD", 0)
    #     continue
    
    # Get volume and costs for calculations
    row_id = row.get("ROW ID #")
    df_row = df[df["ROW ID #"] == row_id].iloc[0] if not df[df["ROW ID #"] == row_id].empty else None
    
    if df_row is not None:
        volume = pd.to_numeric(df_row.get(volume_col), errors='coerce')
        if pd.isna(volume):
            volume = 0
        
        landed_cost = df_row.get(f"{selected_supplier} - R2 - Total landed cost per UOM (USD)", 0)
        if pd.isna(landed_cost):
            landed_cost = 0
        
        extended_cost = landed_cost * volume
        
        # Calculate savings
        fob_usd_col = f"{selected_supplier} - Final USD savings vs baseline"
        landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
        
        try:
            fob_savings = float(df_row.get(fob_usd_col, 0)) if pd.notna(df_row.get(fob_usd_col)) else 0
        except (ValueError, TypeError):
            fob_savings = 0
        
        try:
            landed_savings = float(df_row.get(landed_usd_col, 0)) if pd.notna(df_row.get(landed_usd_col)) else 0
        except (ValueError, TypeError):
            landed_savings = 0
        
        total_fob_savings_usd += fob_savings
        total_landed_savings_usd += landed_savings
        
        # Classify supplier and update metrics
        if selected_supplier == incumbent:
            incumbent_retained += 1
            total_incumbent_volume += volume
            total_landed_cost_incumbent += extended_cost
        else:
            if selected_supplier != '-':
                if selected_supplier in incumbent_suppliers:
                    new_supplier_count += 1
                    total_new_supplier_volume += volume
                    total_landed_cost_new_suppliers += extended_cost
                else:
                    net_new_supplier_count += 1
                    net_new_supplier_list.add(selected_supplier)
                    total_net_new_supplier_volume += volume
                    total_landed_cost_completely_new_suppliers += extended_cost
        
        unique_suppliers.add(selected_supplier)

print(f"Final metrics after all processing:")
print(f"  - Total annual revenue discount: ${total_annual_revenue_discount:,.2f}")
print(f"  - Total landed savings USD: ${total_landed_savings_usd:,.2f}")
print(f"  - Incumbent retained: {incumbent_retained}")
print(f"  - New suppliers: {new_supplier_count}")
print(f"  - Net new suppliers: {net_new_supplier_count}")
print(f"  - Unique suppliers: {len(unique_suppliers)}")

# In output data want to add new column Redundant Suppliers per Product Family.
'''
basically count how many unique selected suppliers are there per product family and add a column named above and add those value for each part.
'''
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

updated_data = [
  {
    "row_id": 9190,
    "Incumbent Supplier": "Waysing",
    "Selected Supplier": "Eaglelite",
    "Part bid on (yes/no)": "Yes",
    "Part #": "SR-800-SP",
    "Extended baseline landed cost USD": 12217,
    "Landed Cost Savings USD": -3484,
    "Landed Extended Cost USD": 15701,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 5347,
    "Incumbent Supplier": "Yuhuan Aigema Copper",
    "Selected Supplier": "Manek Metalcraft",
    "Part bid on (yes/no)": "Yes",
    "Part #": "34732",
    "Extended baseline landed cost USD": 9359,
    "Landed Cost Savings USD": -2187,
    "Landed Extended Cost USD": 11546,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8462,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Meide Group",
    "Part bid on (yes/no)": "Yes",
    "Part #": "942185",
    "Extended baseline landed cost USD": 3141,
    "Landed Cost Savings USD": -192535,
    "Landed Extended Cost USD": 195676,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8664,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Meide Group",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970314",
    "Extended baseline landed cost USD": 1986,
    "Landed Cost Savings USD": -1913,
    "Landed Extended Cost USD": 3898,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 11774,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Meide Group",
    "Part bid on (yes/no)": "Yes",
    "Part #": "9691212CTSLF",
    "Extended baseline landed cost USD": 1190,
    "Landed Cost Savings USD": -1732,
    "Landed Extended Cost USD": 2922,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12440,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Meide Group",
    "Part bid on (yes/no)": "Yes",
    "Part #": "940137X",
    "Extended baseline landed cost USD": 1679,
    "Landed Cost Savings USD": -1799,
    "Landed Extended Cost USD": 3478,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12312,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGE-200",
    "Extended baseline landed cost USD": 3005,
    "Landed Cost Savings USD": -1521,
    "Landed Extended Cost USD": 4526,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12461,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGE-300",
    "Extended baseline landed cost USD": 3211,
    "Landed Cost Savings USD": -2422,
    "Landed Extended Cost USD": 5633,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12462,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGH-600-NBR",
    "Extended baseline landed cost USD": 4222,
    "Landed Cost Savings USD": -7512,
    "Landed Extended Cost USD": 11735,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12514,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGE-600",
    "Extended baseline landed cost USD": 5999,
    "Landed Cost Savings USD": -2803,
    "Landed Extended Cost USD": 8801,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 12515,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGH-800-NBR",
    "Extended baseline landed cost USD": 5081,
    "Landed Cost Savings USD": -9001,
    "Landed Extended Cost USD": 14082,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 13397,
    "Incumbent Supplier": "TANGSHAN HONGYOU CERAMICS",
    "Selected Supplier": "WEFLO",
    "Part bid on (yes/no)": "Yes",
    "Part #": "BFVGH-200-NBR",
    "Extended baseline landed cost USD": 750,
    "Landed Cost Savings USD": -1061,
    "Landed Extended Cost USD": 1811,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8659,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "West Legend-MTD",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970303",
    "Extended baseline landed cost USD": 8143,
    "Landed Cost Savings USD": -2106,
    "Landed Extended Cost USD": 10248,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8661,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "West Legend-MTD",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970305",
    "Extended baseline landed cost USD": 6946,
    "Landed Cost Savings USD": -1271,
    "Landed Extended Cost USD": 8217,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8666,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "West Legend-MTD",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970331",
    "Extended baseline landed cost USD": 14915,
    "Landed Cost Savings USD": -3541,
    "Landed Extended Cost USD": 18456,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8651,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Zhejiang Acme",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970208",
    "Extended baseline landed cost USD": 7749,
    "Landed Cost Savings USD": -1075,
    "Landed Extended Cost USD": 8824,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 8656,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Zhejiang Acme",
    "Part bid on (yes/no)": "Yes",
    "Part #": "970254",
    "Extended baseline landed cost USD": 27459,
    "Landed Cost Savings USD": -1068,
    "Landed Extended Cost USD": 28527,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 9949,
    "Incumbent Supplier": "PT Ever Age",
    "Selected Supplier": "Zhejiang Acme",
    "Part bid on (yes/no)": "Yes",
    "Part #": "44841LF",
    "Extended baseline landed cost USD": 2079,
    "Landed Cost Savings USD": -1473,
    "Landed Extended Cost USD": 3553,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  },
  {
    "row_id": 5893,
    "Incumbent Supplier": "Yuhuan Aigema Copper",
    "Selected Supplier": "ZHEJIANG WANDEKAI",
    "Part bid on (yes/no)": "Yes",
    "Part #": "46931",
    "Extended baseline landed cost USD": 25388,
    "Landed Cost Savings USD": -1008,
    "Landed Extended Cost USD": 26395,
    "Change": "Leave with incumbent at WAPP pricing, 0 savings",
    "Rationale": "Need to keep incumbent anyway due to some of their parts not bid on"
  }
]

# --- Update output_df rows using updated_data JSON ---
for update in updated_data:
    row_id = update["row_id"]
    part_num = update["Part #"]
    selected_supplier = update["Incumbent Supplier"]
    baseline_landed_cost = update["Extended baseline landed cost USD"]
    print(baseline_landed_cost)
    incumbent_supplier = update["Incumbent Supplier"]
    # Find matching row in output_df
    mask = (output_df["ROW ID #"] == row_id) & (output_df["Part #"] == part_num)
    idxs = output_df.index[mask].tolist()
    for idx in idxs:
        # Update supplier fields
        output_df.at[idx, "Selected Supplier"] = selected_supplier
        output_df.at[idx, "Incumbent Supplier"] = incumbent_supplier
        # Update savings and cost fields
        output_df.at[idx, "Landed Cost Savings USD"] = 0
        output_df.at[idx, "Landed Extended Cost USD"] = baseline_landed_cost
        output_df.at[idx, "Change"] = update.get("Change", "")
        output_df.at[idx, "Rationale"] = update.get("Rationale", "")
        # Update all supplier-dependent fields
        pct_col = f"{selected_supplier} - Final % savings vs baseline"
        usd_col = f"{selected_supplier} - Final USD savings vs baseline"
        landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
        landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
        output_df.at[idx, "FOB Savings %"] = output_df.at[idx, pct_col] if pct_col in output_df.columns else "-"
        output_df.at[idx, "FOB Savings USD"] = output_df.at[idx, usd_col] if usd_col in output_df.columns else "-"
        output_df.at[idx, "Landed Cost Savings %"] = output_df.at[idx, landed_pct_col] if landed_pct_col in output_df.columns else "-"
        output_df.at[idx, "Landed Cost Savings USD"] = output_df.at[idx, landed_usd_col] if landed_usd_col in output_df.columns else "-"
        # landed_cost_col = f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"
        # landed_cost = output_df.at[idx, landed_cost_col] if landed_cost_col in output_df.columns else 0
        # try:
        #     landed_cost = float(landed_cost)
        # except:
        #     landed_cost = 0
        # output_df.at[idx, "Landed Extended Cost USD"] = landed_cost * volume
        # Update supplier classification
        output_df.at[idx, "Is Totally New Supplier"] = "Yes" if selected_supplier not in incumbent_suppliers else "No"
        output_df.at[idx, "Part Switched"] = "Yes" if selected_supplier != incumbent_supplier else "No"
        # Update reference data
        ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
        def get_ref_value(col):
            return ref_row[col].values[0] if not ref_row.empty else "-"
        output_df.at[idx, "Standard leadtime - days PO-shipment POL"] = get_ref_value("Standard leadtime - days PO-shipment POL")
        output_df.at[idx, "Retail Packaging"] = get_ref_value("Retail Packaging")
        output_df.at[idx, "Payment term - days and discounts"] = get_ref_value("Payment term - days and discounts")
        output_df.at[idx, "New product introduction"] = get_ref_value("New product introduction")
        output_df.at[idx, "Long term commitment rebate"] = get_ref_value("Long term commitment rebate")
        output_df.at[idx, "Uncompetitive supplier behavior"] = get_ref_value("Uncompetitive supplier behavior")

# Step 5: Move column to index 14
if "Redundant Suppliers per Product Group" in output_df.columns:
    redundancy_col = output_df.pop("Redundant Suppliers per Product Group")
    output_df.insert(14, "Redundant Suppliers per Product Group", redundancy_col)

all_incumbents = set(output_df["Incumbent Supplier"].dropna().unique())

incumbent_rows = output_df[output_df["Selected Supplier"] == output_df["Incumbent Supplier"]]
incumbent_suppliers_unique = set(incumbent_rows["Selected Supplier"].dropna().unique())

new_rows = output_df[output_df["Selected Supplier"] != output_df["Incumbent Supplier"]]
new_suppliers = set(new_rows["Selected Supplier"].dropna().unique())

new_rows = output_df[output_df["Selected Supplier"] != output_df["Incumbent Supplier"]]
new_suppliers_existing = set(
    new_rows["Selected Supplier"].dropna().unique()
).intersection(all_incumbents)

net_new_suppliers = new_suppliers - all_incumbents


#


# --- Summary sheet ---

# # --- Add country column from country_supplier_mapping.csv ---
country_map_path = "scenario_scripts/supplier_country_mapping.csv"
import os
if os.path.exists(country_map_path):
    country_map_df = pd.read_csv(country_map_path)
    # Build lookup for (ROW ID #, Part #, Supplier) -> Country
    country_lookup = {}
    for _, row in country_map_df.iterrows():
        key = (row["ROW ID #"], str(row["Part #"]), str(row["Supplier"]))
        country_lookup[key] = row.get("Country", "-")
    # Prepare country column
    country_col = []
    for idx, df_row in output_df.iterrows():
        key = (df_row["ROW ID #"], str(df_row["Part #"]), str(df_row["Selected Supplier"]))
        country_col.append(country_lookup.get(key, "-"))
    # Insert country column next to Selected Supplier
    sel_idx = output_df.columns.get_loc("Selected Supplier")
    output_df.insert(sel_idx + 1, "Country", country_col)
else:
    print(f"country_supplier_mapping.csv not found, skipping country column.")


total_landed_savings_usd = pd.to_numeric(output_df['Landed Cost Savings USD'], errors='coerce').sum()


total_landed_cost_incumbent = output_df.loc[
    output_df["Incumbent Supplier"] == output_df["Selected Supplier"],
    "Landed Extended Cost USD"
].sum()

total_landed_cost_completely_new_suppliers = output_df.loc[
    ~output_df["Selected Supplier"].isin(incumbent_suppliers),
    "Landed Extended Cost USD"
].sum()

total_landed_cost_new_suppliers = output_df.loc[
    (output_df["Incumbent Supplier"] != output_df["Selected Supplier"]) &
    (output_df["Selected Supplier"].isin(incumbent_suppliers)),
    "Landed Extended Cost USD"
].sum()
summary_data = [
    # ["Total FOB Savings USD", total_fob_savings_usd],
    ["Total Landed Cost Savings USD", total_landed_savings_usd],
    # ["Total Annual Revenue Discount USD", total_annual_revenue_discount],
    # ["Total Cost of no valid suppliers", total_cost_not_awarded],

    ["Total Landed Cost where Incumbent Suppliers Retained", total_landed_cost_incumbent],
    
    ["Total Landed Cost where bid is awarded to New Suppliers", total_landed_cost_new_suppliers],
    
    ["Total Landed Cost where bid is awarded to Completely New Suppliers", total_landed_cost_completely_new_suppliers],
    
    ["Total parts where Incumbent Suppliers Retained", incumbent_retained],
    
    ["Total parts where bid is awarded to New Suppliers", new_supplier_count],
    ["Total parts where bid is awarded to Net New Suppliers", net_new_supplier_count],
    ["Parts not awarded to any supplier", parts_where_no_bids],
    ["", ""],
    ["Totally New Suppliers", len(net_new_supplier_list)],
    ["Total Unique Suppliers", len(unique_suppliers)],

]

output_file = 'scenario3_40 tweaks-3.xlsx'

# --- Write to Excel ---
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario: {round(PERCENT_NEW*100, 2)}% New Supplier, {round((1-PERCENT_NEW)*100, 2)}% Incumbent"
    header_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
    worksheet.merge_range(0, 0, 0, len(list(output_data[0].keys()))-1, scenario_header, header_format)

    # Define formats
    bold_format = workbook.add_format({'bold': True})
    # num_format = workbook.add_format({'num_format': '000,00.00'})
    usd_format = workbook.add_format({'num_format': '$000,00.00'})
    # int_format = workbook.add_format({'num_format': '000,00.00'})

    # Write summary into separate logical blocks:
    summary_row = 2

    # Grouped indices
    cost_metrics = summary_data[0:5]
    supplier_metrics = summary_data[5:15]
    volume_metrics = summary_data[15:]

    # Write cost metrics: Columns A & B
    for i, item in enumerate(cost_metrics):
        worksheet.write(summary_row + i, 0, item[0], bold_format)
        worksheet.write(summary_row + i, 1, item[1], usd_format)

    # Write supplier metrics: Columns D & E
    for i, item in enumerate(supplier_metrics):
        worksheet.write(summary_row + i, 3, item[0], bold_format)
        worksheet.write(summary_row + i, 4, item[1])

    # --- Write total evaluated cost row ---
    total_label_row = summary_row + max(len(cost_metrics), len(supplier_metrics), len(volume_metrics)) + 1
    worksheet.write(total_label_row, 0, "Total Landed Cost Evaluated", bold_format)

    # Formula for summing cost values (adjust B3:B8 if more/less than 6 rows of cost)

    total_cost =  total_landed_savings_usd + total_landed_cost_incumbent + total_landed_cost_new_suppliers + total_landed_cost_completely_new_suppliers
    worksheet.write(total_label_row, 1, total_cost, usd_format)

    # Write output table
    df_output = output_df
    df_output.to_excel(writer, sheet_name="Sheet1", startrow=13, index=False)

# --- Timer ---
elapsed_time = time.time() - start_time
print(f"\n✅ Done. Output written to '{output_file}'")
print(f"⏱ Time taken: {elapsed_time:.2f} seconds")