"""
Computation Engine for ICAI Revenue Recognition (POCM).
Implements all 9 steps of Percentage Completion Method per ICAI Guidance Note (Revised 2012).
Closing Inventory computed as per Illustration in the Guidance Note.
"""
import pandas as pd


def parse_uploaded_file(uploaded_file):
    """Parse uploaded Excel file and return structured data dict."""
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    except Exception as e:
        return None, f"Error reading file: {e}"

    required = ['PROJECT_MASTER', 'UNIT_WISE_DATA', 'COST_INCURRED']
    missing = [s for s in required if s not in xls.sheet_names]
    if missing:
        return None, f"Missing sheets: {', '.join(missing)}"

    # Parse PROJECT_MASTER (vertical layout: Field in A, Value in B)
    df_pm = pd.read_excel(xls, 'PROJECT_MASTER', skiprows=2, header=0)
    pm = {}
    if len(df_pm.columns) >= 2:
        for _, row in df_pm.iterrows():
            key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            val = row.iloc[1] if pd.notna(row.iloc[1]) else ''
            pm[key] = val

    # Parse UNIT_WISE_DATA
    df_units = pd.read_excel(xls, 'UNIT_WISE_DATA', skiprows=2, header=0)
    # Drop rows that are entirely empty
    df_units = df_units.dropna(how='all').reset_index(drop=True)
    col_map = {}
    for c in df_units.columns:
        cl = str(c).lower().strip()
        if 'unit' in cl and 'no' in cl: col_map[c] = 'unit_no'
        elif 'tower' in cl or 'phase' in cl: col_map[c] = 'tower'
        elif 'saleable' in cl or ('area' in cl and 'sq' in cl): col_map[c] = 'saleable_area'
        elif 'agreement' in cl and 'val' in cl: col_map[c] = 'agreement_value'
        elif 'amount' in cl and 'realised' in cl: col_map[c] = 'amount_realised'
        elif 'realised' in cl and '%' in cl: col_map[c] = 'pct_realised'
        elif '10%' in cl or 'threshold' in cl: col_map[c] = 'threshold_met'
        elif 'eligible' in cl: col_map[c] = 'eligible'
        elif 'status' in cl: col_map[c] = 'status'
    df_units = df_units.rename(columns=col_map)

    # Filter out blank template skeleton rows:
    # A real unit must have either a non-zero agreement_value or a non-blank unit_no.
    if 'agreement_value' in df_units.columns:
        df_units['agreement_value'] = pd.to_numeric(df_units['agreement_value'], errors='coerce').fillna(0)
        has_agreement = df_units['agreement_value'] > 0
    else:
        has_agreement = pd.Series([False] * len(df_units))
    if 'unit_no' in df_units.columns:
        has_unit_no = df_units['unit_no'].notna() & (df_units['unit_no'].astype(str).str.strip() != '') & \
                      (df_units['unit_no'].astype(str).str.strip().str.lower() != 'nan')
    else:
        has_unit_no = pd.Series([False] * len(df_units))
    df_units = df_units[has_agreement | has_unit_no].reset_index(drop=True)

    # Parse COST_INCURRED
    df_cost = pd.read_excel(xls, 'COST_INCURRED', skiprows=2, header=0)
    costs = {}
    if len(df_cost.columns) >= 2:
        for _, row in df_cost.iterrows():
            key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            val = row.iloc[1] if pd.notna(row.iloc[1]) else 0
            costs[key] = val

    return {'project_master': pm, 'units': df_units, 'costs': costs}, None


def _safe_num(val, default=0):
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _get_pm_val(pm, partial_key):
    for k, v in pm.items():
        if partial_key.lower() in k.lower():
            return v
    return 0


def validate_thresholds(data):
    """Validate ICAI Para 5.3 thresholds (construction >=25%, area sold >=25%)."""
    pm = data['project_master']
    units = data['units']
    costs = data['costs']

    total_saleable = _safe_num(_get_pm_val(pm, 'Total Saleable'))
    est_construction = _safe_num(_get_pm_val(pm, 'Estimated Construction'))
    est_land = _safe_num(_get_pm_val(pm, 'Estimated Land'))
    est_other = _safe_num(_get_pm_val(pm, 'Estimated Other'))
    est_total = est_land + est_construction + est_other

    construction_incurred = _safe_num(_get_pm_val(costs, 'Construction Cost'))
    land_incurred = _safe_num(_get_pm_val(costs, 'Land Cost'))
    other_incurred = _safe_num(_get_pm_val(costs, 'Other'))
    total_incurred = land_incurred + construction_incurred + other_incurred

    construction_pct = (construction_incurred / est_construction * 100) if est_construction > 0 else 0

    if 'status' in units.columns and 'saleable_area' in units.columns:
        active = units[units['status'].astype(str).str.strip().str.lower() == 'active']
        area_sold = _safe_num(active['saleable_area'].sum())
    else:
        area_sold = 0
    area_sold_pct = (area_sold / total_saleable * 100) if total_saleable > 0 else 0

    return {
        'construction_pct': construction_pct,
        'construction_pass': construction_pct >= 25,
        'area_sold': area_sold,
        'area_sold_pct': area_sold_pct,
        'area_pass': area_sold_pct >= 25,
        'all_pass': construction_pct >= 25 and area_sold_pct >= 25,
        'est_total': est_total,
        'est_construction': est_construction,
        'est_land': est_land,
        'est_other': est_other,
        'total_incurred': total_incurred,
        'land_incurred': land_incurred,
        'construction_incurred': construction_incurred,
        'other_incurred': other_incurred,
        'total_saleable': total_saleable,
    }


def compute_revenue_recognition(data):
    """
    Run full POCM computation per ICAI Guidance Note illustration.
    Closing Inventory split into:
      - Inventory of Unsold Units = (Unsold Area / Total Area) x Total Cost Incurred
      - WIP of Sold Units = (Sold Area / Total Area) x Total Cost Incurred - Cost of Revenue
    """
    pm = data['project_master']
    units = data['units']
    costs = data['costs']

    thresholds = validate_thresholds(data)
    est_total = thresholds['est_total']
    total_incurred = thresholds['total_incurred']
    total_saleable = thresholds['total_saleable']

    # STEP 1 - Stage of Completion
    stage_of_completion = (total_incurred / est_total) if est_total > 0 else 0

    # Mark eligible contracts
    if 'agreement_value' in units.columns and 'amount_realised' in units.columns:
        units = units.copy()
        units['agreement_value'] = pd.to_numeric(units.get('agreement_value', 0), errors='coerce').fillna(0)
        units['amount_realised'] = pd.to_numeric(units.get('amount_realised', 0), errors='coerce').fillna(0)
        units['saleable_area'] = pd.to_numeric(units.get('saleable_area', 0), errors='coerce').fillna(0)
        units['pct_realised_calc'] = units.apply(
            lambda r: r['amount_realised'] / r['agreement_value'] if r['agreement_value'] > 0 else 0, axis=1)
        units['threshold_10_met'] = units['pct_realised_calc'] >= 0.10
        status_col = units.get('status', pd.Series(['Active'] * len(units)))
        units['is_active'] = status_col.astype(str).str.strip().str.lower() == 'active'
        units['is_eligible'] = units['threshold_10_met'] & units['is_active']
    else:
        units = pd.DataFrame()

    # STEP 2 - Threshold check (only construction >= 25% and area sold >= 25%)
    all_thresholds_met = thresholds['all_pass']

    # STEP 3 - Eligible Revenue & Area
    if len(units) > 0:
        eligible_revenue = units.loc[units['is_eligible'], 'agreement_value'].sum()
        total_amount_realised = units.loc[units['is_eligible'], 'amount_realised'].sum()
        area_sold = units.loc[units['is_active'], 'saleable_area'].sum()
        total_agreement_value = units.loc[units['is_active'], 'agreement_value'].sum()
        total_realised_all = units.loc[units['is_active'], 'amount_realised'].sum()
    else:
        eligible_revenue = 0
        total_amount_realised = 0
        area_sold = 0
        total_agreement_value = 0
        total_realised_all = 0

    revenue_to_recognise = min(stage_of_completion * eligible_revenue, eligible_revenue) if all_thresholds_met else 0

    # STEP 4 - Cost of Revenue (per ICAI Illustration)
    # Cost of Revenue = Stage of Completion x (Area Sold / Total Area) x Total Estimated Cost
    # Which simplifies to: (Area Sold / Total Area) x Total Cost Incurred
    sold_area_ratio = (area_sold / total_saleable) if total_saleable > 0 else 0
    unsold_area = total_saleable - area_sold

    # Per ICAI Guidance Note illustration:
    cost_of_revenue = stage_of_completion * sold_area_ratio * est_total

    # STEP 5 - Closing Inventory Breakdown (per ICAI Illustration)
    # Inventory of Unsold Units = (Unsold Area / Total Area) x Total Cost Incurred
    unsold_area_ratio = (unsold_area / total_saleable) if total_saleable > 0 else 0
    inventory_unsold_units = unsold_area_ratio * total_incurred

    # WIP of Sold Units = (Sold Area / Total Area) x Total Cost Incurred - Cost of Revenue
    cost_incurred_sold = sold_area_ratio * total_incurred
    wip_sold_units = cost_incurred_sold - cost_of_revenue

    # Total Closing Inventory / WIP
    total_closing_inventory = inventory_unsold_units + wip_sold_units

    # Cross-check: Total Closing Inventory = Total Cost Incurred - Cost of Revenue
    # (This must always balance)

    # STEP 6 - Profit
    profit = revenue_to_recognise - cost_of_revenue

    # STEP 7 - Unbilled Revenue
    unbilled_revenue = max(0, revenue_to_recognise - total_amount_realised)

    # STEP 8 - Advances
    advances = max(0, total_amount_realised - revenue_to_recognise)

    # STEP 9 - Expected Loss
    expected_loss = max(0, est_total - eligible_revenue) if est_total > eligible_revenue and eligible_revenue > 0 else 0

    # Pending Realisation = Total Agreement Value (Active) - Total Amount Realised (Active)
    pending_realisation = total_agreement_value - total_realised_all

    # Warnings
    warnings = []
    if not thresholds['construction_pass']:
        warnings.append("Construction completion ({:.1f}%) is below 25% threshold [Para 5.3(b)]".format(
            thresholds['construction_pct']))
    if not thresholds['area_pass']:
        warnings.append("Area sold ({:.1f}%) is below 25% threshold [Para 5.3(c)]".format(
            thresholds['area_sold_pct']))
    if len(units) > 0:
        # Only count genuinely filled rows (agreement_value > 0)
        filled = units[units['agreement_value'] > 0] if 'agreement_value' in units.columns else units
        non_eligible = filled[~filled['is_eligible']]
        if len(non_eligible) > 0:
            warnings.append("{} unit(s) not eligible (10% realisation not met or cancelled)".format(len(non_eligible)))
        cancelled = filled[~filled['is_active']]
        if len(cancelled) > 0:
            warnings.append("{} unit(s) cancelled - revenue reversed".format(len(cancelled)))
    if expected_loss > 0:
        warnings.append("EXPECTED LOSS of Rs.{:,.0f} must be recognised immediately [Para 5.7]".format(expected_loss))
    if est_total > 0 and total_incurred > est_total:
        warnings.append("Cost incurred (Rs.{:,.0f}) exceeds estimates (Rs.{:,.0f}) - COST OVERRUN".format(
            total_incurred, est_total))

    est_revenue = _safe_num(_get_pm_val(pm, 'Total Estimated Project Revenue'))
    project_name = _get_pm_val(pm, 'Project Name')

    return {
        'project_name': project_name if project_name else 'Unnamed Project',
        'stage_of_completion': stage_of_completion,
        'all_thresholds_met': all_thresholds_met,
        'thresholds': thresholds,
        'eligible_revenue': eligible_revenue,
        'revenue_to_recognise': revenue_to_recognise,
        'cost_of_revenue': cost_of_revenue,
        'profit': profit,
        'inventory_unsold_units': inventory_unsold_units,
        'wip_sold_units': wip_sold_units,
        'total_closing_inventory': total_closing_inventory,
        'cost_incurred_sold': cost_incurred_sold,
        'unsold_area': unsold_area,
        'unbilled_revenue': unbilled_revenue,
        'advances': advances,
        'expected_loss': expected_loss,
        'total_amount_realised': total_amount_realised,
        'total_realised_all': total_realised_all,
        'area_sold': area_sold,
        'total_saleable': total_saleable,
        'est_total': est_total,
        'est_revenue': est_revenue,
        'total_incurred': total_incurred,
        'total_agreement_value': total_agreement_value,
        'pending_realisation': pending_realisation,
        'units': units,
        'warnings': warnings,
    }
