"""Test the updated computation engine with sample data including unsold units."""
import sys
sys.path.insert(0, r"c:\Users\Jimi Modi\Desktop\AI Class\AICA Level 2\CAPSTONE PROJECT")
from engine import parse_uploaded_file, compute_revenue_recognition

data, err = parse_uploaded_file(
    r"c:\Users\Jimi Modi\Desktop\AI Class\AICA Level 2\CAPSTONE PROJECT\sample_filled.xlsx"
)
if err:
    print("Parse error:", err)
    sys.exit(1)

r = compute_revenue_recognition(data)

print("=" * 70)
print("PROJECT: {}".format(r['project_name']))
print("=" * 70)

print("\n--- PROJECT OVERVIEW ---")
print("Total Saleable Area: {:,.0f} sq.ft".format(r['total_saleable']))
print("Total Area Sold: {:,.0f} sq.ft".format(r['area_sold']))
print("Unsold Area: {:,.0f} sq.ft".format(r['unsold_area']))
print("Sale Consideration (Agreements): {:,.0f}".format(r['total_agreement_value']))
print("Total Estimated Revenue: {:,.0f}".format(r['est_revenue']))
print("Pending Realisation: {:,.0f}".format(r['pending_realisation']))

print("\n--- THRESHOLDS ---")
print("Construction %: {:.1f}% (Pass: {})".format(r['thresholds']['construction_pct'], r['thresholds']['construction_pass']))
print("Area Sold %: {:.1f}% (Pass: {})".format(r['thresholds']['area_sold_pct'], r['thresholds']['area_pass']))
print("All Thresholds Met: {}".format(r['all_thresholds_met']))

print("\n--- REVENUE & PROFIT ---")
print("Stage of Completion: {:.2%}".format(r['stage_of_completion']))
print("Eligible Revenue: {:,.0f}".format(r['eligible_revenue']))
print("Revenue Recognised: {:,.0f}".format(r['revenue_to_recognise']))
print("Cost of Revenue: {:,.0f}".format(r['cost_of_revenue']))
print("Profit: {:,.0f}".format(r['profit']))

print("\n--- CLOSING INVENTORY / WIP (Per ICAI Illustration) ---")
print("A. Inventory of Unsold Units: {:,.0f}".format(r['inventory_unsold_units']))
print("   = ({:,.0f} / {:,.0f}) x {:,.0f}".format(r['unsold_area'], r['total_saleable'], r['total_incurred']))
print("B. WIP of Sold Units: {:,.0f}".format(r['wip_sold_units']))
print("   = ({:,.0f} / {:,.0f}) x {:,.0f} - {:,.0f}".format(
    r['area_sold'], r['total_saleable'], r['total_incurred'], r['cost_of_revenue']))
print("Total Closing Inventory: {:,.0f}".format(r['total_closing_inventory']))
print("Cross-check (Incurred - Cost of Rev): {:,.0f}".format(r['total_incurred'] - r['cost_of_revenue']))

print("\n--- OTHER ---")
print("Unbilled Revenue: {:,.0f}".format(r['unbilled_revenue']))
print("Advances: {:,.0f}".format(r['advances']))
print("Expected Loss: {:,.0f}".format(r['expected_loss']))

print("\n--- WARNINGS ---")
for w in r['warnings']:
    print("  " + w)

print("\n--- UNIT SUMMARY ---")
units = r['units']
if len(units) > 0:
    eligible = int(units['is_eligible'].sum())
    active = int(units['is_active'].sum())
    total = len(units)
    unsold_units = units[(units['agreement_value'] == 0) | (units['agreement_value'].isna())]
    print("Total: {}, Active: {}, Eligible: {}, Non-Eligible: {}".format(total, active, eligible, total - eligible))
    print("Unsold units (no agreement): {}".format(len(unsold_units)))
    if len(unsold_units) > 0:
        for _, u in unsold_units.iterrows():
            print("  {} - {} ({} sq.ft)".format(u.get('unit_no','?'), u.get('tower','?'), u.get('saleable_area',0)))

print("\n" + "=" * 70)
# Verify inventory balances
balance_check = abs(r['total_closing_inventory'] - (r['total_incurred'] - r['cost_of_revenue']))
print("INVENTORY BALANCE CHECK: {} (diff: {:.2f})".format(
    "PASS" if balance_check < 0.01 else "FAIL", balance_check))
print("=" * 70)
