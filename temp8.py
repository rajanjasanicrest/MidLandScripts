import pandas as pd

# Load CSV file
df = pd.read_csv("new/Bidsheet Master Consolidate Landed With Updated Tariff.csv")

# Initialize counters and totals
incumbent_gained_part = 0
incumbent_landed_cost = 0.0

missing_incumbent_quote_parts = 0
missing_incumbent_landed_cost = 0.0

# Iterate over DataFrame rows
for idx, row in df.iterrows():

    if row.get('Valid Supplier') == 0:
        continue
    selected_supplier = row["Final Minimum Bid Landed Supplier"]
    incumbent_supplier = row["Normalized incumbent supplier"]
    volume = row.get("Annual Volume (per UOM)", 0)

    # Step 1: Selected supplier is the incumbent
    if selected_supplier == incumbent_supplier:
        cost_col = f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"
        # cost_col = f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"
        cost = row.get(cost_col, 0)
        if pd.notna(cost) and pd.notna(volume):
            incumbent_gained_part += 1
            incumbent_landed_cost += cost * volume

    # Step 2: Incumbent column missing or has non-positive value
    inc_col = f"{incumbent_supplier} - R2 - Total landed cost per UOM (USD)"
    cost = row.get(inc_col, None)
    if inc_col not in df.columns or pd.isna(cost) or cost <= 0:
        cost_col = f"{selected_supplier} - R2 - Total landed cost per UOM (USD)"
        landed_cost = row.get(cost_col, 0)
        if pd.notna(landed_cost) and pd.notna(volume):
            missing_incumbent_quote_parts += 1
            missing_incumbent_landed_cost += landed_cost * volume

# Final Output
print(f"✅ Incumbent Gained Parts: {incumbent_gained_part}")
print(f"💰 Incumbent Landed Cost: ${incumbent_landed_cost:,.2f}")
print(f"❌ Missing Incumbent Quote Parts: {missing_incumbent_quote_parts}")
print(f"💰 Landed Cost (Missing Incumbent): ${missing_incumbent_landed_cost:,.2f}")
