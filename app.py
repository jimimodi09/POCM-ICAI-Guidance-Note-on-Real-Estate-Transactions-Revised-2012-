"""
ICAI Revenue Recognition Tool - Main Streamlit Application
Implements POCM per ICAI Guidance Note (Revised 2012)
"""
import streamlit as st
import pandas as pd
from template_generator import generate_template
from engine import parse_uploaded_file, compute_revenue_recognition
from output_generator import generate_output_excel

# --- Page Config ---
st.set_page_config(
    page_title="ICAI Revenue Recognition Tool - POCM",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
* { font-family: 'Inter', sans-serif; }
.main .block-container { padding: 1.5rem 2rem; max-width: 1200px; }
.app-header {
    background: linear-gradient(135deg, #1B3A5C 0%, #2C5F8A 50%, #1B3A5C 100%);
    color: white; padding: 1.8rem 2.5rem; border-radius: 12px;
    margin-bottom: 1.5rem; box-shadow: 0 4px 20px rgba(27,58,92,0.25);
}
.app-header h1 { margin: 0 0 0.3rem 0; font-size: 1.7rem; font-weight: 700; letter-spacing: -0.5px; }
.app-header p { margin: 0; opacity: 0.85; font-size: 0.9rem; font-weight: 300; }
.kpi-card {
    background: #fff; border-radius: 10px; padding: 1.1rem 1.3rem;
    border-left: 4px solid #2C5F8A; box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    transition: transform 0.2s; height: 100%; min-height: 85px;
}
.kpi-card:hover { transform: translateY(-2px); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
.kpi-label { font-size: 0.7rem; color: #666; text-transform: uppercase; letter-spacing: 0.5px;
             font-weight: 600; margin-bottom: 0.3rem; }
.kpi-value { font-size: 1.35rem; font-weight: 700; color: #1B3A5C; }
.kpi-value.green { color: #1B5E20; }
.kpi-value.red { color: #B71C1C; }
.kpi-card.accent { border-left-color: #F57F17; }
.kpi-card.info { border-left-color: #0277BD; }
.threshold-pass { background: #E8F5E9; color: #1B5E20; padding: 0.5rem 1rem;
                  border-radius: 8px; font-weight: 600; text-align: center;
                  border: 1px solid #A5D6A7; }
.threshold-fail { background: #FFEBEE; color: #B71C1C; padding: 0.5rem 1rem;
                  border-radius: 8px; font-weight: 600; text-align: center;
                  border: 1px solid #EF9A9A; }
.section-hdr {
    font-size: 1.1rem; font-weight: 700; color: #1B3A5C;
    border-bottom: 2px solid #2C5F8A; padding-bottom: 0.4rem;
    margin: 1.2rem 0 0.8rem 0;
}
.warn-box {
    background: #FFF8E1; border-left: 4px solid #F57F17; padding: 0.8rem 1.2rem;
    border-radius: 6px; margin-bottom: 0.5rem; font-size: 0.88rem;
}
.warn-box.critical { background: #FFEBEE; border-left-color: #B71C1C; }
[data-testid="stSidebar"] { background: #F5F7FA; }
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown("""
<div class="app-header">
    <h1>🏗️ ICAI Revenue Recognition Tool</h1>
    <p>Percentage of Completion Method (POCM) — As per ICAI Guidance Note on Real Estate Transactions (Revised 2012)</p>
</div>
""", unsafe_allow_html=True)

# --- Sidebar ---
with st.sidebar:
    st.markdown("### 📋 How to Use")
    st.markdown("""
    **Step 1:** Download the Excel template below.\n
    **Step 2:** Fill in your project data in the template.\n
    **Step 3:** Upload the filled template.\n
    **Step 4:** View auto-computed results and download reports.
    """)
    st.divider()
    st.markdown("### ⬇️ Download Template")
    template_buf = generate_template()
    st.download_button(
        label="📥 Download Excel Template",
        data=template_buf,
        file_name="ICAI_POCM_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )
    st.divider()
    st.markdown("### 📤 Upload Filled Template")
    uploaded = st.file_uploader("Choose your filled Excel file", type=['xlsx', 'xls'],
                                label_visibility="collapsed")
    st.divider()
    st.markdown("""
    <div style="font-size: 0.75rem; color: #888; text-align: center;">
    Built for CA Professionals<br>
    ICAI Guidance Note (Revised 2012)<br>
    Para 5.1 – 5.9 Compliance
    </div>
    """, unsafe_allow_html=True)


# --- Helpers ---
def fmt_inr(val):
    try:
        val = float(val)
    except (ValueError, TypeError):
        return "₹0"
    if val < 0:
        return "(₹{:,.0f})".format(abs(val))
    return "₹{:,.0f}".format(val)


def kpi_card(label, value, color="", extra_cls=""):
    cls = " {}".format(color) if color else ""
    ecls = " {}".format(extra_cls) if extra_cls else ""
    return '<div class="kpi-card{}"><div class="kpi-label">{}</div><div class="kpi-value{}">{}</div></div>'.format(
        ecls, label, cls, value)


def threshold_badge(label, value_pct, passed, threshold="25%"):
    cls = "threshold-pass" if passed else "threshold-fail"
    icon = "✅" if passed else "❌"
    return '<div class="{}">{} {}<br><b>{}</b> (Threshold: {})</div>'.format(cls, icon, label, value_pct, threshold)


# --- Main Content ---
if uploaded is None:
    st.markdown('<div class="section-hdr">📌 Getting Started</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        #### 1️⃣ Download Template
        Download the pre-formatted Excel template from the sidebar.
        It contains 4 sheets with embedded formulas.
        """)
    with col2:
        st.markdown("""
        #### 2️⃣ Fill Your Data
        Enter project details, unit-wise sales data,
        and costs incurred in the respective sheets.
        """)
    with col3:
        st.markdown("""
        #### 3️⃣ Upload & Analyse
        Upload the filled template and get instant
        audit-ready revenue recognition computations.
        """)

    st.divider()
    st.markdown('<div class="section-hdr">📖 ICAI Compliance Framework</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        **Revenue Recognition Conditions (Para 5.3):**
        - Construction completion ≥ 25%
        - Total area sold ≥ 25%
        - Minimum 10% realisation per contract
        """)
    with c2:
        st.markdown("""
        **Outputs Generated:**
        - Revenue Recognition Working Sheet
        - Profit / Loss Computation
        - Closing Inventory (Unsold + WIP Sold)
        - Journal Entries & Disclosure Notes
        - Compliance Report with Para References
        """)

else:
    # Parse
    data, error = parse_uploaded_file(uploaded)
    if error:
        st.error("❌ **File Error:** {}".format(error))
        st.stop()

    results = compute_revenue_recognition(data)
    r = results
    t = r['thresholds']

    # === TABS ===
    tab_dash, tab_rev, tab_profit, tab_inv, tab_je, tab_disc, tab_comply, tab_units = st.tabs([
        "📊 Dashboard", "💰 Revenue Working", "📈 Profit Summary",
        "🏗️ Inventory & WIP", "📝 Journal Entries", "📋 Disclosures",
        "✅ Compliance", "👥 Unit Analysis"
    ])

    # --- TAB: Dashboard ---
    with tab_dash:
        # ROW 1: Project Overview KPIs
        st.markdown('<div class="section-hdr">Project Overview</div>', unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown(kpi_card("Total Saleable Area", "{:,.0f} sq.ft".format(r['total_saleable']), extra_cls="info"),
                        unsafe_allow_html=True)
        with c2:
            st.markdown(kpi_card("Total Area Sold", "{:,.0f} sq.ft".format(r['area_sold']), extra_cls="info"),
                        unsafe_allow_html=True)
        with c3:
            st.markdown(kpi_card("Sale Consideration (Agreements)", fmt_inr(r['total_agreement_value']), extra_cls="info"),
                        unsafe_allow_html=True)
        with c4:
            st.markdown(kpi_card("Total Estimated Revenue", fmt_inr(r['est_revenue']), extra_cls="info"),
                        unsafe_allow_html=True)
        with c5:
            st.markdown(kpi_card("Pending Realisation", fmt_inr(r['pending_realisation']), extra_cls="accent"),
                        unsafe_allow_html=True)

        # ROW 2: Revenue & Profit KPIs
        st.markdown('<div class="section-hdr">Revenue Recognition Summary</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(kpi_card("Revenue Recognised", fmt_inr(r['revenue_to_recognise'])), unsafe_allow_html=True)
        with c2:
            st.markdown(kpi_card("Cost of Revenue (P&L)", fmt_inr(r['cost_of_revenue'])), unsafe_allow_html=True)
        with c3:
            color = "green" if r['profit'] >= 0 else "red"
            st.markdown(kpi_card("Profit / (Loss)", fmt_inr(r['profit']), color), unsafe_allow_html=True)
        with c4:
            st.markdown(kpi_card("Stage of Completion", "{:.2%}".format(r['stage_of_completion'])),
                        unsafe_allow_html=True)

        # ROW 3: Inventory & Other
        st.markdown("")
        if not r['all_thresholds_met']:
            # PCM not triggered — all costs on Balance Sheet as one Inventory/WIP block
            c5, c6, c7, c8 = st.columns(4)
            with c5:
                st.markdown(
                    kpi_card("Total Closing Inventory / WIP (Balance Sheet)",
                             fmt_inr(r['total_closing_inventory']), extra_cls="accent"),
                    unsafe_allow_html=True)
            with c6:
                st.markdown(kpi_card("Unbilled Revenue", fmt_inr(r['unbilled_revenue'])), unsafe_allow_html=True)
            with c7:
                adv_color = "amber" if r['advances'] > 0 else ""
                st.markdown(kpi_card("Advances from Customers", fmt_inr(r['advances']), adv_color),
                            unsafe_allow_html=True)
            with c8:
                color = "red" if r['expected_loss'] > 0 else "green"
                lbl = fmt_inr(r['expected_loss']) if r['expected_loss'] > 0 else "NIL"
                st.markdown(kpi_card("Expected Loss", lbl, color), unsafe_allow_html=True)
            # Info banner explaining why
            st.markdown(
                '<div style="background:#FFF3E0;border-left:4px solid #E65100;padding:0.7rem 1.2rem;'
                'border-radius:6px;font-size:0.85rem;margin-top:0.5rem;">'
                '📦 <b>All construction costs are carried as Inventory / WIP</b> — '
                'PCM conditions (Para 5.3) not yet satisfied. '
                'Revenue = NIL &nbsp;|&nbsp; Cost of Revenue = NIL &nbsp;|&nbsp; Profit = NIL</div>',
                unsafe_allow_html=True)
        else:
            c5, c6, c7, c8, c9 = st.columns(5)
            with c5:
                st.markdown(kpi_card("Inventory - Unsold Units", fmt_inr(r['inventory_unsold_units'])),
                            unsafe_allow_html=True)
            with c6:
                st.markdown(kpi_card("WIP - Sold Units", fmt_inr(r['wip_sold_units'])), unsafe_allow_html=True)
            with c7:
                st.markdown(kpi_card("Unbilled Revenue", fmt_inr(r['unbilled_revenue'])), unsafe_allow_html=True)
            with c8:
                adv_color = "amber" if r['advances'] > 0 else ""
                st.markdown(kpi_card("Advances from Customers", fmt_inr(r['advances']), adv_color),
                            unsafe_allow_html=True)
            with c9:
                color = "red" if r['expected_loss'] > 0 else "green"
                lbl = fmt_inr(r['expected_loss']) if r['expected_loss'] > 0 else "NIL"
                st.markdown(kpi_card("Expected Loss", lbl, color), unsafe_allow_html=True)


        # Threshold Status
        st.markdown('<div class="section-hdr">Threshold Validation (Para 5.3) — All Three Must Be Satisfied</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(threshold_badge("Para 5.3(b) — Construction & Dev. Cost",
                                        "{:.1f}%".format(t['construction_pct']),
                                        t['construction_pass']), unsafe_allow_html=True)
        with c2:
            st.markdown(threshold_badge("Para 5.3(c) — Area Secured by Agreements",
                                        "{:.1f}%".format(t['area_sold_pct']),
                                        t['area_pass']), unsafe_allow_html=True)
        with c3:
            realisation_label = "At least one contract ≥ 10% realised" if t['realisation_pass'] else "No contract has ≥ 10% realisation"
            st.markdown(threshold_badge("Para 5.3(d) — Minimum 10% Realisation",
                                        realisation_label,
                                        t['realisation_pass'],
                                        threshold="Any contract ≥ 10%"), unsafe_allow_html=True)

        # Warnings
        if r['warnings']:
            st.markdown('<div class="section-hdr">⚠️ Warning Flags</div>', unsafe_allow_html=True)
            for w in r['warnings']:
                cls = "critical" if "EXPECTED LOSS" in w or "COST OVERRUN" in w else ""
                st.markdown('<div class="warn-box {}">{}</div>'.format(cls, w), unsafe_allow_html=True)

        # Download
        st.markdown("")
        output_buf = generate_output_excel(results)
        st.download_button(
            label="📥 Download Audit-Ready Output Excel",
            data=output_buf,
            file_name="ICAI_POCM_Output_{}.xlsx".format(r['project_name'].replace(' ', '_')),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

    # --- TAB: Revenue Working ---
    with tab_rev:
        st.markdown('<div class="section-hdr">Revenue Recognition Working Sheet</div>', unsafe_allow_html=True)
        rev_data = {
            'Particulars': [
                'Total Estimated Project Cost', 'Total Cost Incurred to Date',
                'Stage of Completion (%)', '',
                'Construction Completion %', 'Area Sold %',
                'All Thresholds Met?', '',
                'Eligible Revenue (Sum of Eligible Contracts)',
                'Revenue to Recognise (Stage x Eligible Revenue, capped)',
            ],
            'Value': [
                fmt_inr(r['est_total']), fmt_inr(r['total_incurred']),
                '{:.2%}'.format(r['stage_of_completion']), '',
                '{:.2f}%'.format(t['construction_pct']),
                '{:.2f}%'.format(t['area_sold_pct']),
                'YES ✅' if r['all_thresholds_met'] else 'NO ❌', '',
                fmt_inr(r['eligible_revenue']),
                fmt_inr(r['revenue_to_recognise']),
            ],
            'ICAI Reference': [
                'Para 5.2', 'Para 5.2', 'Para 5.6', '',
                'Para 5.3(b)', 'Para 5.3(c)', 'Para 5.3', '',
                'Para 5.3(d)', 'Para 5.1',
            ]
        }
        st.dataframe(pd.DataFrame(rev_data), use_container_width=True, hide_index=True)

    # --- TAB: Profit Summary ---
    with tab_profit:
        st.markdown('<div class="section-hdr">Profit / Loss Computation</div>', unsafe_allow_html=True)
        profit_data = {
            'Particulars': ['Revenue Recognised', 'Less: Cost of Revenue', 'Profit / (Loss) for the Period'],
            'Amount': [fmt_inr(r['revenue_to_recognise']), fmt_inr(r['cost_of_revenue']), fmt_inr(r['profit'])]
        }
        st.dataframe(pd.DataFrame(profit_data), use_container_width=True, hide_index=True)

        st.markdown("")
        st.markdown("**Cost Bifurcation:**")
        cost_data = {
            'Cost Component': ['Land Cost', 'Construction Cost', 'Other / Admin Cost', 'Total'],
            'Estimated': [fmt_inr(t['est_land']), fmt_inr(t['est_construction']),
                              fmt_inr(t['est_other']), fmt_inr(t['est_total'])],
            'Incurred': [fmt_inr(t['land_incurred']), fmt_inr(t['construction_incurred']),
                             fmt_inr(t['other_incurred']), fmt_inr(t['total_incurred'])],
        }
        st.dataframe(pd.DataFrame(cost_data), use_container_width=True, hide_index=True)

    # --- TAB: Inventory & WIP ---
    with tab_inv:
        st.markdown('<div class="section-hdr">Closing Inventory / WIP</div>',
                    unsafe_allow_html=True)

        if not r['all_thresholds_met']:
            # PCM conditions NOT met — ENTIRE cost goes to Balance Sheet, nil to P&L
            st.warning(
                "⚠️ **PCM conditions (Para 5.3) not met** — Revenue and Cost of Revenue are NIL. "
                "All construction costs are carried as **Inventory / Work-in-Progress** on the Balance Sheet "
                "until the conditions are satisfied."
            )
            inv_data = {
                'Particulars': [
                    'Total Cost Incurred to Date',
                    'Less: Cost of Revenue charged to P&L',
                    '',
                    'Closing Inventory / WIP (entire cost — Balance Sheet)',
                    '  [No unsold/sold split applied — PCM not yet triggered]',
                ],
                'Amount': [
                    fmt_inr(r['total_incurred']),
                    'NIL',
                    '',
                    fmt_inr(r['total_closing_inventory']),
                    '',
                ]
            }
        else:
            inv_data = {
                'Particulars': [
                    'Total Cost Incurred', 'Less: Cost of Revenue (P&L)', '',
                    'A. Inventory of Unsold Units',
                    '   Unsold Area (sq.ft): {:,.0f}'.format(r['unsold_area']),
                    '   = (Unsold Area / Total Area) x Total Cost Incurred', '',
                    'B. WIP of Sold Units (cost not yet recognised)',
                    '   = (Sold Area / Total Area) x Cost Incurred - Cost of Revenue', '',
                    'Total Closing Inventory / WIP (A + B)',
                    'Cross-check: Total Incurred - Cost of Revenue',
                ],
                'Amount': [
                    fmt_inr(r['total_incurred']), fmt_inr(r['cost_of_revenue']), '',
                    fmt_inr(r['inventory_unsold_units']), '', '', '',
                    fmt_inr(r['wip_sold_units']), '', '',
                    fmt_inr(r['total_closing_inventory']),
                    fmt_inr(r['total_incurred'] - r['cost_of_revenue']),
                ]
            }
        st.dataframe(pd.DataFrame(inv_data), use_container_width=True, hide_index=True)

        if r['expected_loss'] > 0:
            st.markdown('<div class="section-hdr">Expected Loss Working (Para 5.7)</div>', unsafe_allow_html=True)
            loss_data = {
                'Particulars': ['Total Estimated Project Cost', 'Eligible Revenue', 'Expected Loss'],
                'Amount': [fmt_inr(r['est_total']), fmt_inr(r['eligible_revenue']), fmt_inr(r['expected_loss'])]
            }
            st.dataframe(pd.DataFrame(loss_data), use_container_width=True, hide_index=True)
            st.error("🔴 **Expected Loss of {} must be recognised immediately per Para 5.7**".format(
                fmt_inr(r['expected_loss'])))

    # --- TAB: Journal Entries ---
    with tab_je:
        st.markdown('<div class="section-hdr">Journal Entry Sheet</div>', unsafe_allow_html=True)
        entries = []

        if r['revenue_to_recognise'] > 0:
            # ── Entry 1: Revenue Recognition ──────────────────────────────────
            entries.append(('**1. Revenue Recognition  [Para 5.1 / 5.4]**', '', ''))

            # Cash/Bank debit = min(amount realised, revenue recognised)
            # (cannot debit more cash than the revenue being recognised)
            cash_debit = min(r['total_amount_realised'], r['revenue_to_recognise'])
            if cash_debit > 0:
                entries.append(('    Cash / Bank A/c  (Amount Realised from Eligible Contracts)',
                                fmt_inr(cash_debit), ''))

            # Unbilled Revenue = Revenue recognised - Cash collected (positive remainder)
            unbilled = max(0, r['revenue_to_recognise'] - r['total_amount_realised'])
            if unbilled > 0:
                entries.append(('    Unbilled Revenue A/c  (Revenue > Cash Received)',
                                fmt_inr(unbilled), ''))

            # Credit: full revenue recognised
            entries.append(('        To  Revenue from Operations A/c',
                            '', fmt_inr(r['revenue_to_recognise'])))
            entries.append(('', '', ''))

            # ── Entry 2: Cost Recognition ──────────────────────────────────────
            entries.append(('**2. Cost Recognition  [Para 5.6]**', '', ''))
            entries.append(('    Cost of Revenue A/c',
                            fmt_inr(r['cost_of_revenue']), ''))
            entries.append(('        To  Work-in-Progress / Inventory A/c',
                            '', fmt_inr(r['cost_of_revenue'])))

            # ── Entry 3: Advances refund if cash > revenue (Para 9 disclosure) ──
            # Excess cash received beyond revenue recognised → Advances from Customers
            excess_cash = max(0, r['total_amount_realised'] - r['revenue_to_recognise'])
            if excess_cash > 0:
                entries.append(('', '', ''))
                entries.append(('**3. Excess Cash Received  [Para 9 — Advances from Customers]**', '', ''))
                entries.append(('    Cash / Bank A/c  (Remaining collections — not yet recognised)',
                                fmt_inr(excess_cash), ''))
                entries.append(('        To  Advances from Customers A/c',
                                '', fmt_inr(excess_cash)))

        else:
            # ── PCM NOT triggered: all cash collected treated as Advance ──────
            if r['total_amount_realised'] > 0:
                entries.append(('**1. Collections from Customers  [PCM not yet triggered — Para 5.3]**', '', ''))
                entries.append(('    Cash / Bank A/c  (Amount collected from customers)',
                                fmt_inr(r['total_amount_realised']), ''))
                entries.append(('        To  Advances from Customers A/c',
                                '', fmt_inr(r['total_amount_realised'])))
                entries.append(('', '', ''))
                entries.append(('    [Note: Revenue recognition deferred. All costs carried as Inventory / WIP.]',
                                '', ''))

        # ── Expected Loss provision (always, if applicable) ───────────────────
        if r['expected_loss'] > 0:
            entry_num = '3' if r['revenue_to_recognise'] > 0 and r['total_amount_realised'] <= r['revenue_to_recognise'] else \
                        '4' if r['revenue_to_recognise'] > 0 else '2'
            entries.append(('', '', ''))
            entries.append(('**{}. Expected Loss Provision  [Para 5.7]**'.format(entry_num), '', ''))
            entries.append(('    Loss on Real Estate Project A/c',
                            fmt_inr(r['expected_loss']), ''))
            entries.append(('        To  Provision for Expected Loss A/c',
                            '', fmt_inr(r['expected_loss'])))

        if entries:
            je_df = pd.DataFrame(entries, columns=['Account', 'Debit (₹)', 'Credit (₹)'])
            st.dataframe(je_df, use_container_width=True, hide_index=True)
        else:
            st.info("No journal entries — revenue recognition criteria not met and no cash collected.")



    # --- TAB: Disclosures ---
    with tab_disc:
        st.markdown('<div class="section-hdr">Disclosure Note as per Para 9</div>', unsafe_allow_html=True)
        st.markdown("""
**Project:** {project}

**Method of Revenue Recognition:** Percentage of Completion Method (POCM)
as per ICAI Guidance Note on Accounting for Real Estate Transactions (Revised 2012)

| Disclosure Item | Amount |
|---|---|
| Stage of Completion | {soc} |
| Revenue Recognised | {rev} |
| Cost of Revenue | {cost} |
| Profit / (Loss) | {profit} |
| Inventory - Unsold Units | {inv_unsold} |
| WIP - Sold Units | {wip_sold} |
| Total Closing Inventory / WIP | {total_inv} |
| Unbilled Revenue | {unbilled} |
| Advances from Customers | {advances} |
""".format(
            project=r['project_name'],
            soc='{:.2%}'.format(r['stage_of_completion']),
            rev=fmt_inr(r['revenue_to_recognise']),
            cost=fmt_inr(r['cost_of_revenue']),
            profit=fmt_inr(r['profit']),
            inv_unsold=fmt_inr(r['inventory_unsold_units']),
            wip_sold=fmt_inr(r['wip_sold_units']),
            total_inv=fmt_inr(r['total_closing_inventory']),
            unbilled=fmt_inr(r['unbilled_revenue']),
            advances=fmt_inr(r['advances']),
        ))
        if r['expected_loss'] > 0:
            st.markdown("| Expected Loss Recognised | {} |".format(fmt_inr(r['expected_loss'])))

    # --- TAB: Compliance ---
    with tab_comply:
        st.markdown('<div class="section-hdr">Compliance Report - ICAI Guidance Note (Revised 2012)</div>',
                    unsafe_allow_html=True)
        checks = [
            ('Para 5.3(b)', 'Construction & dev. cost incurred >= 25% of total estimated cost', t['construction_pass']),
            ('Para 5.3(c)', 'Area secured by agreements >= 25% of total saleable area', t['area_pass']),
            ('Para 5.3(d)', 'At least 10% of agreement value realised per active contract', t['realisation_pass']),
            ('Para 5.4', 'Revenue recognised using POCM', r['all_thresholds_met']),
            ('Para 5.6', 'Cost incurred method used for stage of completion', True),
            ('Para 5.7', 'Expected loss recognised in full if applicable', True),
        ]
        comp_df = pd.DataFrame(checks, columns=['ICAI Para', 'Requirement', 'Compliant'])
        comp_df['Status'] = comp_df['Compliant'].apply(lambda x: '✅ COMPLIANT' if x else '❌ NON-COMPLIANT')
        st.dataframe(comp_df[['ICAI Para', 'Requirement', 'Status']], use_container_width=True, hide_index=True)

    # --- TAB: Unit Analysis ---
    with tab_units:
        st.markdown('<div class="section-hdr">Unit-wise Contract Analysis</div>', unsafe_allow_html=True)
        if isinstance(r['units'], pd.DataFrame) and len(r['units']) > 0:
            display_cols = []
            for c in ['unit_no', 'tower', 'saleable_area', 'agreement_value',
                      'amount_realised', 'pct_realised_calc', 'threshold_10_met',
                      'is_active', 'is_eligible']:
                if c in r['units'].columns:
                    display_cols.append(c)
            df_display = r['units'][display_cols].copy()
            rename = {
                'unit_no': 'Unit No', 'tower': 'Tower', 'saleable_area': 'Area (Sq.ft)',
                'agreement_value': 'Agreement Value', 'amount_realised': 'Amt Realised',
                'pct_realised_calc': '% Realised', 'threshold_10_met': '10% Met',
                'is_active': 'Active', 'is_eligible': 'Eligible'
            }
            df_display = df_display.rename(columns=rename)
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            eligible_count = r['units']['is_eligible'].sum() if 'is_eligible' in r['units'].columns else 0
            total_count = len(r['units'])
            st.markdown("**Total Units:** {} | **Eligible:** {} | **Non-Eligible:** {}".format(
                total_count, int(eligible_count), total_count - int(eligible_count)))
        else:
            st.info("No unit data found in the uploaded file.")
