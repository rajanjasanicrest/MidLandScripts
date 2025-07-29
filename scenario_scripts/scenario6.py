import pandas as pd
from tqdm import tqdm
import time

# --- Constants ---
CODA = "Coda"
MEIDE = "Meide Group"
CODA_FOB_LIMIT = 2_000_000
MEIDE_FOB_LIMIT = 17_500_000
BONUS_SAVINGS = {
    CODA: 0.075 * CODA_FOB_LIMIT,  # 7.5%
    MEIDE: 0.05 * MEIDE_FOB_LIMIT  # 5%
}

incumbent_col = "Normalized incumbent supplier"
qty_col = "Annual Volume (per UOM)"

# --- Load files ---
start_time = time.time()
df = pd.read_csv("new/Bidsheet Master Consolidate Landed.csv")
output_reference_df = pd.read_csv("new/outout-reference.csv")
print(f"✅ Loaded {len(df)} rows\n")

# --- Step 1: AS-IS BASELINE (incumbent if present, else lowest bidder) ---
output_data = []
coda_candidates = []
meide_candidates = []
supplier_fob_spend = {CODA: 0, MEIDE: 0}
supplier_savings_usd = {CODA: 0, MEIDE: 0}

print("Building AS-IS assignment and strategic reallocation pool...\n")
for idx, row in tqdm(df.iterrows(), total=len(df), desc="Analyzing"):
    incumbent = row.get(incumbent_col)
    min_supplier = row.get("Final Minimum Bid Landed Supplier")
    selected_supplier = incumbent if pd.notna(incumbent) else min_supplier
    if pd.isna(selected_supplier):
        selected_supplier = "No valid suppliers"

    fob_col = f"{selected_supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
    fob_savings_pct = row.get(f"{selected_supplier} - Final % savings vs baseline", "-")
    fob_savings_usd = row.get(f"{selected_supplier} - Final USD savings vs baseline", 0)
    qty = row.get(qty_col, 1)
    fob_spend = row.get(fob_col, 0) * qty if pd.notna(row.get(fob_col)) else 0

    # Check CODA eligibility
    coda_col = f"{CODA} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
    if CODA != incumbent and coda_col in row and pd.notna(row[coda_col]):
        coda_fob = row[coda_col] * qty
        coda_saving = row.get(f"{CODA} - Final USD savings vs baseline", 0)
        try:
            coda_saving = float(coda_saving)
        except:
            coda_saving = 0
        coda_candidates.append((idx, coda_saving, coda_fob))

    # Check MEIDE eligibility
    meide_col = f"{MEIDE} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"
    if MEIDE != incumbent and meide_col in row and pd.notna(row[meide_col]):
        meide_fob = row[meide_col] * qty
        meide_saving = row.get(f"{MEIDE} - Final USD savings vs baseline", 0)
        try:
            meide_saving = float(meide_saving)
        except:
            meide_saving = 0
        meide_candidates.append((idx, meide_saving, meide_fob))

    output_data.append({
        "idx": idx,
        "Selected Supplier": selected_supplier,
        "FOB Savings %": fob_savings_pct,
        "FOB Savings USD": fob_savings_usd,
        "Reason": "As-Is (incumbent or lowest)"
    })

# --- Step 2: Allocate to CODA ---
coda_candidates.sort(key=lambda x: x[1], reverse=True)  # sort by savings desc
for idx, usd_saving, fob in coda_candidates:
    if supplier_fob_spend[CODA] + fob <= CODA_FOB_LIMIT:
        supplier_fob_spend[CODA] += fob
        supplier_savings_usd[CODA] += usd_saving
        output_data[idx].update({
            "Selected Supplier": CODA,
            "FOB Savings %": df.at[idx, f"{CODA} - Final % savings vs baseline"],
            "FOB Savings USD": df.at[idx, f"{CODA} - Final USD savings vs baseline"],
            "Reason": "Strategic CODA allocation"
        })

# --- Step 3: Allocate to MEIDE ---
meide_candidates.sort(key=lambda x: x[1], reverse=True)
for idx, usd_saving, fob in meide_candidates:
    if supplier_fob_spend[MEIDE] + fob <= MEIDE_FOB_LIMIT:
        supplier_fob_spend[MEIDE] += fob
        supplier_savings_usd[MEIDE] += usd_saving
        output_data[idx].update({
            "Selected Supplier": MEIDE,
            "FOB Savings %": df.at[idx, f"{MEIDE} - Final % savings vs baseline"],
            "FOB Savings USD": df.at[idx, f"{MEIDE} - Final USD savings vs baseline"],
            "Reason": "Strategic MEIDE allocation"
        })

# --- Step 4: Final output ---
final_rows = []
for row in output_data:
    idx = row["idx"]
    selected_supplier = row["Selected Supplier"]
    landed_pct_col = f"{selected_supplier} - Final Landed % savings vs baseline"
    landed_usd_col = f"{selected_supplier} - Final Landed USD savings vs baseline"
    ref_row = output_reference_df[output_reference_df['Reference'] == selected_supplier]
    def get_ref(col): return ref_row[col].values[0] if not ref_row.empty else "-"

    final_rows.append({
        "ROW ID #": df.at[idx, "ROW ID #"],
        "Division": df.at[idx, "Division"],
        "Part #": df.at[idx, "Part #"],
        "Item Description": df.at[idx, "Item Description"],
        "Product Group": df.at[idx, "Product Group"],
        "Part Family": df.at[idx, "Part Family"],
        "Part Family": df.at[idx, "Normalized incumbent supplier"],
        "Selected Supplier": selected_supplier,
        "FOB Savings %": row.get("FOB Savings %", "-"),
        "FOB Savings USD": row.get("FOB Savings USD", "-"),
        "Landed Cost Savings %": df.at[idx, landed_pct_col] if landed_pct_col in df.columns else "-",
        "Landed Cost Savings USD": df.at[idx, landed_usd_col] if landed_usd_col in df.columns else "-",
        "Reason": row["Reason"],
        "Part Switched": "Yes" if selected_supplier != df.at[idx, incumbent_col] else "No",
        "Uncompetitive supplier behavior": get_ref("Uncompetitive supplier behavior"),
        "Standard leadtime - days PO-shipment POL": get_ref("Standard leadtime - days PO-shipment POL"),
        "Retail Packaging": get_ref("Retail Packaging"),
        "Payment term - days and discounts": get_ref("Payment term - days and discounts"),
        "New product introduction": get_ref("New product introduction"),
        "Long term commitment rebate": get_ref("Long term commitment rebate"),
    })

# --- Apply bonuses ---
supplier_savings_usd[CODA] += BONUS_SAVINGS[CODA]
supplier_savings_usd[MEIDE] += BONUS_SAVINGS[MEIDE]

# --- Output to Excel ---
pd.DataFrame(final_rows).to_excel("Scenario_6_Strategic_Reallocation.xlsx", index=False)

elapsed = time.time() - start_time
print(f"\n✅ Scenario 6 complete. Output saved to 'Scenario_6_Strategic_Reallocation.xlsx'")
print(f"⏱ Time taken: {elapsed:.2f} sec")
print(f"💸 Coda allocated: ${supplier_fob_spend[CODA]:,.2f} | Total Savings: ${supplier_savings_usd[CODA]:,.2f}")
print(f"💸 Meide allocated: ${supplier_fob_spend[MEIDE]:,.2f} | Total Savings: ${supplier_savings_usd[MEIDE]:,.2f}")
