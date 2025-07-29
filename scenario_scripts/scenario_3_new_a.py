import pandas as pd

# Constants
TOTAL_BASELINE_COST = 49_783_721.6480354
TARGET_SAVINGS = 6_250_000

# Load data
df = pd.read_csv("new/Bidsheet Master Consolidate Landed2.csv")

# Helpers
def cost_col(supplier):
    return f"{supplier} - R2 - Total landed cost per UOM (USD)"
def save_col(supplier):
    return f"{supplier} - Final Landed USD savings vs baseline"

# Step 1: Calculate locked savings and forced savings, and keep track of forced new supplier landed costs
locked_savings = 0.0
forced_savings = 0.0
incumbent_locked_landed_cost = 0.0
forced_new_landed_cost = 0.0

for _, row in df.iterrows():
    if row.get("Valid Supplier", 0) == 0:
        continue
    
    inc = row.get("Normalized incumbent supplier")
    new = row.get("Final Minimum Bid Landed Supplier")
    vol = row.get("Annual Volume (per UOM)", 0)
    
    if pd.isna(new) or vol <= 0:
        continue
    
    # Locked parts (incumbent is min supplier)
    if new == inc:
        s = row.get(save_col(inc), 0) or 0
        locked_savings += s
        c = row.get(cost_col(inc), 0) or 0
        incumbent_locked_landed_cost += c * vol
    
    # Forced new (no incumbent bid)
    inc_cost = row.get(cost_col(inc), None) if pd.notna(inc) else None
    if inc_cost is None or pd.isna(inc_cost) or inc_cost <= 0:
        s = row.get(save_col(new), 0) or 0
        forced_savings += s
        c = row.get(cost_col(new), 0) or 0
        forced_new_landed_cost += c * vol

print(f"Locked savings (incumbent min): ${locked_savings:,.2f}")
print(f"Forced new supplier savings:    ${forced_savings:,.2f}")
print(f"Locked incumbent landed cost:   ${incumbent_locked_landed_cost:,.2f}")
print(f"Forced new supplier landed cost:${forced_new_landed_cost:,.2f}")

# Step 2: Build list of switchable parts (where incumbent != new supplier)
switchable_parts = []

for _, row in df.iterrows():
    if row.get("Valid Supplier", 0) == 0:
        continue
    
    inc = row.get("Normalized incumbent supplier")
    new = row.get("Final Minimum Bid Landed Supplier")
    vol = row.get("Annual Volume (per UOM)", 0)
    
    if pd.notna(inc) and pd.notna(new) and inc != new and vol > 0:
        inc_savings = row.get(save_col(inc), 0) or 0
        new_savings = row.get(save_col(new), 0) or 0
        inc_cost = row.get(cost_col(inc), 0) or 0
        new_cost = row.get(cost_col(new), 0) or 0
        
        # Extra savings from switching incumbent -> new supplier
        extra_savings = new_savings - inc_savings
        
        baseline_cost = inc_cost * vol
        new_landed_cost = new_cost * vol
        
        switchable_parts.append({
            "index": row.name,
            "inc_savings": inc_savings,
            "new_savings": new_savings,
            "extra_savings": extra_savings,
            "baseline_cost": baseline_cost,
            "new_landed_cost": new_landed_cost
        })

temp_switchable_landed_cost = sum(p["new_landed_cost"] for p in switchable_parts)

print(f"Switchable parts count: {len(switchable_parts)}")
print(f"Switchable Landed Cost: ${sum(p['new_landed_cost'] for p in switchable_parts):,.2f}")

# Step 3: Calculate initial total savings assuming all switchable parts assigned to incumbent
initial_switchable_inc_savings = sum(p["inc_savings"] for p in switchable_parts)
total_initial_savings = locked_savings + forced_savings + initial_switchable_inc_savings

print(f"Initial incumbent savings on switchable parts: ${initial_switchable_inc_savings:,.2f}")
print(f"Total savings if no switch on switchable parts: ${total_initial_savings:,.2f}")

# Step 4: Calculate how much more savings needed to hit target
remaining_needed = TARGET_SAVINGS - total_initial_savings

# If already reached or exceeded target, no switching needed
if remaining_needed <= 0:
    pct_switch = (forced_new_landed_cost) / TOTAL_BASELINE_COST * 100
    print("\nTarget already reached with locked and forced savings.")
    print(f"Percentage switched needed: {pct_switch:.2f}%")
    exit()

# Step 5: Sort switchable parts by descending extra savings (switch benefit)
switchable_parts.sort(key=lambda x: x["extra_savings"], reverse=True)

print(switchable_parts[:10])

# Step 6: Greedy assign switches until target reached
current_savings = total_initial_savings
switched_landed_cost = forced_new_landed_cost
switched_count = 0

for part in switchable_parts:

    if current_savings >= TARGET_SAVINGS:
        break

    else:
        current_savings = current_savings + part["extra_savings"]
        switched_landed_cost += part["new_landed_cost"]
        switched_count += 1

        temp_switchable_landed_cost -= part["new_landed_cost"]

# Step 7: Calculate percentage switched of baseline spend
pct_switch = switched_landed_cost / (temp_switchable_landed_cost + switched_landed_cost + incumbent_locked_landed_cost) * 100

print("\n--- Final Results ---")
print(f"Target savings:               ${TARGET_SAVINGS:,.2f}")
print(f"Locked savings (incumbent):   ${locked_savings:,.2f}")
print(f"Forced new supplier savings:  ${forced_savings:,.2f}")
print(f"Initial incumbent savings on switchable parts: ${initial_switchable_inc_savings:,.2f}")
print(f"Current total savings achieved:        ${current_savings:,.2f}")
print(f"Total new supplier landed cost (forced + switched): ${switched_landed_cost:,.2f}")
print(f"Locked incumbent landed cost:   ${incumbent_locked_landed_cost:,.2f}")
print(f"Number of switched parts:       {switched_count}")
print(f"Percentage of baseline spend switched to new suppliers: {pct_switch:.2f}%")
