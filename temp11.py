import pandas as pd

# --- Inputs ---
allocation_file = "scenario3_40 tweaks-6.xlsx"    
updated_bids_file = "new/Bidsheet Master Consolidate Landed v_t_2.csv"  

def safe_float(val, precision=4):
    try:
        return round(float(val), precision)
    except (ValueError, TypeError):
        return val   # return original if not convertible


# --- Load ---
alloc_df = pd.read_excel(allocation_file, sheet_name="Sheet1", skiprows=13)
bids_df = pd.read_csv(updated_bids_file)

alloc_df["ROW ID #"] = alloc_df["ROW ID #"].astype(str)
bids_df["ROW ID #"] = bids_df["ROW ID #"].astype(str)

volume_col = "Annual Volume (per UOM)"

# --- Update with new prices ---
for idx, row in alloc_df.iterrows():
    part = row["ROW ID #"]
    supplier = row["Selected Supplier"]

    if supplier == "-" or pd.isna(supplier):
        print('hi there')
        continue

    landed_col = f"{supplier} - R2 - Total landed cost per UOM (USD)"
    fob_col = f"{supplier} - R2 - Total Cost Per UOM FOB Port of Origin/Departure (USD)"

    bid_row = bids_df[bids_df["ROW ID #"] == part]
    if bid_row.empty:
        continue

    if landed_col not in bids_df.columns:
        # alloc_df.at[idx, "FOB Savings USD"] = 0
        continue
        # alloc_df.at[idx, "FOB Savings %"] = 0

        # alloc_df.at[idx, "Landed Cost Savings USD"] = 0
        # alloc_df.at[idx, "Landed Cost Savings %"] = 0

        # alloc_df.at[idx, "Landed Extended Cost USD"] = safe_float(bid_row.iloc[0]["Landed Extended Cost USD"])
    else:
        landed_price = safe_float(bid_row.iloc[0][landed_col])
        if pd.isna(landed_price) or landed_price <= 0:
            continue
        else:
         
            alloc_df.at[idx, "FOB Savings USD"] = safe_float(bid_row.iloc[0][f"{supplier} - Final USD savings vs baseline"])
            alloc_df.at[idx, "FOB Savings %"] = safe_float(bid_row.iloc[0][f"{supplier} - Final % savings vs baseline"])

            alloc_df.at[idx, "Landed CostxSavings USD"] = safe_float(bid_row.iloc[0][f"{supplier} - Final Landed USD savings vs baseline"])
            alloc_df.at[idx, "Landed Cost Savings %"] = safe_float(bid_row.iloc[0][f"{supplier} - Final Landed % savings vs baseline"])

            alloc_df.at[idx, "Landed Extended Cost USD"] = landed_price * row[volume_col]

# --- Recalculate summary ---
total_landed_savings = pd.to_numeric(alloc_df["Landed Cost Savings USD"], errors="coerce").sum()
total_landed_cost_incumbent = alloc_df.loc[
    alloc_df["Incumbent Supplier"] == alloc_df["Selected Supplier"], "Landed Extended Cost USD"
].sum()
total_landed_cost_completely_new = alloc_df.loc[
    ~alloc_df["Selected Supplier"].isin(alloc_df["Incumbent Supplier"].unique()), "Landed Extended Cost USD"
].sum()
total_landed_cost_new_suppliers = alloc_df.loc[
    (alloc_df["Selected Supplier"] != alloc_df["Incumbent Supplier"]) &
    (alloc_df["Selected Supplier"].isin(alloc_df["Incumbent Supplier"].unique())),
    "Landed Extended Cost USD"
].sum()

incumbent_retained = (alloc_df["Selected Supplier"] == alloc_df["Incumbent Supplier"]).sum()
new_supplier_count = ((alloc_df["Selected Supplier"] != alloc_df["Incumbent Supplier"]) &
                      (alloc_df["Selected Supplier"].isin(alloc_df["Incumbent Supplier"].unique()))).sum()
net_new_supplier_count = ((alloc_df["Selected Supplier"] != alloc_df["Incumbent Supplier"]) &
                          (~alloc_df["Selected Supplier"].isin(alloc_df["Incumbent Supplier"].unique()))).sum()
parts_no_bids = (alloc_df["Selected Supplier"] == "-").sum()

incumbent_suppliers = alloc_df['Incumbent Supplier'].unique()

total_landed_savings_usd = pd.to_numeric(alloc_df['Landed Cost Savings USD'], errors='coerce').sum()

total_landed_cost_completely_new_suppliers = alloc_df.loc[
    ~alloc_df["Selected Supplier"].isin(incumbent_suppliers),
    "Landed Extended Cost USD"
].sum()

parts_where_no_bids = 0

summary_data = [
    ["Total Landed Cost Savings USD", total_landed_savings_usd],
    ["Total Landed Cost where Incumbent Suppliers Retained", total_landed_cost_incumbent],
    ["Total Landed Cost where bid is awarded to New Suppliers", total_landed_cost_new_suppliers],
    ["Total Landed Cost where bid is awarded to Completely New Suppliers", total_landed_cost_completely_new_suppliers],
    ["Total parts where Incumbent Suppliers Retained", incumbent_retained],
    ["Total parts where bid is awarded to New Suppliers", new_supplier_count],
    ["Total parts where bid is awarded to Net New Suppliers", net_new_supplier_count],
    ["Parts not awarded to any supplier", 0],
    ["", ""],
    ["Totally New Suppliers", 9],
    ["Total Unique Suppliers", 62],
]

# --- Save with summary ---
updated_file = "scenario3_40 tweaks-6_UPDATED.xlsx"
with pd.ExcelWriter(updated_file, engine="xlsxwriter") as writer:
    PERCENT_NEW = 0.40
    workbook = writer.book
    worksheet = workbook.add_worksheet("Sheet1")
    writer.sheets["Sheet1"] = worksheet

    # Scenario header
    scenario_header = f"Scenario: {round(PERCENT_NEW*100, 2)}% New Supplier, {round((1-PERCENT_NEW)*100, 2)}% Incumbent"
    header_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'left'})
    # worksheet.merge_range(0, 0, 0, len(list(alloc_df[0].keys()))-1, scenario_header, header_format)

    # Define formats
    bold_format = workbook.add_format({'bold': True})
    # num_format = workbook.add_format({'num_format': '000,00.00'})
    usd_format = workbook.add_format({'num_format': '$000,00.00'})
    # int_format = workbook.add_format({'num_format': '000,00.00'})

    # Write summary into separate logical blocks:
    summary_row = 2

    # Grouped indices
    cost_metrics = summary_data[0:4]
    supplier_metrics = summary_data[4:15]

    # Write cost metrics: Columns A & B
    for i, item in enumerate(cost_metrics):
        worksheet.write(summary_row + i, 0, item[0], bold_format)
        worksheet.write(summary_row + i, 1, item[1], usd_format)

    # Write supplier metrics: Columns D & E
    for i, item in enumerate(supplier_metrics):
        worksheet.write(summary_row + i, 3, item[0], bold_format)
        worksheet.write(summary_row + i, 4, item[1])

    # --- Write total evaluated cost row ---
    total_label_row = summary_row + max(len(cost_metrics), len(supplier_metrics)) + 1
    worksheet.write(total_label_row, 0, "Total Landed Cost Evaluated", bold_format)

    # Formula for summing cost values (adjust B3:B8 if more/less than 6 rows of cost)

    total_cost =  total_landed_savings_usd + total_landed_cost_incumbent + total_landed_cost_new_suppliers + total_landed_cost_completely_new_suppliers
    worksheet.write(total_label_row, 1, total_cost, usd_format)

    alloc_df.to_excel(writer, sheet_name="Sheet1", startrow=13, index=False)

print(f"✅ Updated file with summary written: {updated_file}")
