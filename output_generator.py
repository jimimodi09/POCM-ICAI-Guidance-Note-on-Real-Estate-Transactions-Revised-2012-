"""
Output Excel Generator - Produces professional audit-ready output workbook.
Sheets: 1) Dashboard  2) Revenue Working  3) Profit Summary
        4) Inventory & WIP  5) Journal Entries  6) Disclosure Note
        7) Unit Analysis  8) Compliance Report  9) Warnings (if any)
Journal Entries match exactly what is displayed in the Streamlit tab.
"""
import io
import xlsxwriter
import pandas as pd
from datetime import datetime


# ── Indian number formatting helper ──────────────────────────────────────────
def _inr(val):
    try:
        val = float(val)
    except (ValueError, TypeError):
        return "₹0"
    if val < 0:
        return "(₹{:,.0f})".format(abs(val))
    return "₹{:,.0f}".format(val)


def _pct(val):
    try:
        return "{:.2f}%".format(float(val) * 100)
    except (ValueError, TypeError):
        return "0.00%"


# ── Format palette ────────────────────────────────────────────────────────────
def _make_formats(wb):
    f = {}

    # Titles / headers
    f['main_title'] = wb.add_format({
        'bold': True, 'font_size': 16, 'font_color': '#FFFFFF',
        'bg_color': '#1B3A5C', 'align': 'center', 'valign': 'vcenter',
        'border': 0
    })
    f['sub_title'] = wb.add_format({
        'italic': True, 'font_size': 9, 'font_color': '#888888',
        'align': 'center', 'valign': 'vcenter'
    })
    f['sheet_title'] = wb.add_format({
        'bold': True, 'font_size': 13, 'font_color': '#1B3A5C',
        'bg_color': '#EAF0F7', 'border': 1, 'align': 'left', 'valign': 'vcenter'
    })
    f['col_hdr'] = wb.add_format({
        'bold': True, 'bg_color': '#1B3A5C', 'font_color': '#FFFFFF',
        'border': 1, 'font_size': 10, 'align': 'center', 'valign': 'vcenter',
        'text_wrap': True
    })
    f['col_hdr_accent'] = wb.add_format({
        'bold': True, 'bg_color': '#2C5F8A', 'font_color': '#FFFFFF',
        'border': 1, 'font_size': 10, 'align': 'center', 'valign': 'vcenter'
    })
    f['section_hdr'] = wb.add_format({
        'bold': True, 'bg_color': '#2C5F8A', 'font_color': '#FFFFFF',
        'border': 1, 'font_size': 11, 'align': 'left', 'valign': 'vcenter'
    })

    # Labels & values
    f['lbl'] = wb.add_format({
        'bg_color': '#EAF0F7', 'border': 1, 'font_size': 10, 'bold': True,
        'valign': 'vcenter'
    })
    f['lbl_indent'] = wb.add_format({
        'bg_color': '#F7FAFD', 'border': 1, 'font_size': 10,
        'valign': 'vcenter', 'indent': 2
    })
    f['val_num'] = wb.add_format({
        'border': 1, 'num_format': '₹#,##,##0', 'font_size': 10,
        'align': 'right', 'valign': 'vcenter'
    })
    f['val_pct'] = wb.add_format({
        'border': 1, 'num_format': '0.00%', 'font_size': 10,
        'align': 'right', 'valign': 'vcenter'
    })
    f['val_txt'] = wb.add_format({
        'border': 1, 'font_size': 10, 'align': 'left', 'valign': 'vcenter'
    })
    f['val_area'] = wb.add_format({
        'border': 1, 'num_format': '#,##,##0', 'font_size': 10,
        'align': 'right', 'valign': 'vcenter'
    })

    # Status badges
    f['pass'] = wb.add_format({
        'bold': True, 'bg_color': '#E8F5E9', 'font_color': '#1B5E20',
        'border': 1, 'font_size': 10, 'align': 'center', 'valign': 'vcenter'
    })
    f['fail'] = wb.add_format({
        'bold': True, 'bg_color': '#FFEBEE', 'font_color': '#B71C1C',
        'border': 1, 'font_size': 10, 'align': 'center', 'valign': 'vcenter'
    })
    f['warn_lbl'] = wb.add_format({
        'bg_color': '#FFF3E0', 'font_color': '#E65100', 'border': 1,
        'font_size': 10, 'bold': True
    })
    f['warn_val'] = wb.add_format({
        'bg_color': '#FFF3E0', 'font_color': '#E65100', 'border': 1,
        'font_size': 10, 'num_format': '₹#,##,##0'
    })

    # Dashboard KPI boxes
    f['kpi_lbl'] = wb.add_format({
        'bold': True, 'bg_color': '#1B3A5C', 'font_color': '#B0C4DE',
        'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1,
        'text_wrap': True
    })
    f['kpi_val'] = wb.add_format({
        'bold': True, 'bg_color': '#FFFFFF', 'font_color': '#1B3A5C',
        'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f['kpi_val_green'] = wb.add_format({
        'bold': True, 'bg_color': '#E8F5E9', 'font_color': '#1B5E20',
        'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f['kpi_val_red'] = wb.add_format({
        'bold': True, 'bg_color': '#FFEBEE', 'font_color': '#B71C1C',
        'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    f['kpi_val_amber'] = wb.add_format({
        'bold': True, 'bg_color': '#FFF8E1', 'font_color': '#F57F17',
        'font_size': 13, 'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    # Journal entry
    f['je_section'] = wb.add_format({
        'bold': True, 'bg_color': '#1B3A5C', 'font_color': '#FFFFFF',
        'border': 1, 'font_size': 11
    })
    f['je_debit'] = wb.add_format({
        'bg_color': '#F7FAFD', 'border': 1, 'num_format': '₹#,##,##0',
        'font_size': 10, 'align': 'right', 'valign': 'vcenter'
    })
    f['je_credit'] = wb.add_format({
        'bg_color': '#F0FFF0', 'border': 1, 'num_format': '₹#,##,##0',
        'font_size': 10, 'italic': True, 'align': 'right', 'valign': 'vcenter'
    })
    f['je_acc_dr'] = wb.add_format({
        'bg_color': '#F7FAFD', 'border': 1, 'font_size': 10,
        'valign': 'vcenter', 'indent': 1
    })
    f['je_acc_cr'] = wb.add_format({
        'bg_color': '#F0FFF0', 'border': 1, 'font_size': 10,
        'italic': True, 'valign': 'vcenter', 'indent': 3
    })
    f['je_total_lbl'] = wb.add_format({
        'bold': True, 'bg_color': '#EAF0F7', 'border': 1, 'font_size': 10,
        'align': 'right'
    })
    f['je_total_val'] = wb.add_format({
        'bold': True, 'bg_color': '#EAF0F7', 'border': 1,
        'num_format': '₹#,##,##0', 'font_size': 10, 'align': 'right'
    })

    # Unit analysis
    f['unit_hdr'] = wb.add_format({
        'bold': True, 'bg_color': '#1B3A5C', 'font_color': '#FFFFFF',
        'border': 1, 'font_size': 10, 'align': 'center', 'text_wrap': True
    })
    f['unit_yes'] = wb.add_format({
        'bg_color': '#E8F5E9', 'font_color': '#1B5E20', 'border': 1,
        'font_size': 9, 'align': 'center', 'bold': True
    })
    f['unit_no'] = wb.add_format({
        'bg_color': '#FFEBEE', 'font_color': '#B71C1C', 'border': 1,
        'font_size': 9, 'align': 'center', 'bold': True
    })
    f['unit_num'] = wb.add_format({
        'border': 1, 'num_format': '#,##,##0', 'font_size': 9,
        'align': 'right'
    })
    f['unit_pct'] = wb.add_format({
        'border': 1, 'num_format': '0.00%', 'font_size': 9, 'align': 'right'
    })
    f['unit_txt'] = wb.add_format({
        'border': 1, 'font_size': 9, 'align': 'center'
    })
    f['unit_txt_l'] = wb.add_format({
        'border': 1, 'font_size': 9, 'align': 'left'
    })
    f['unit_alt'] = wb.add_format({
        'bg_color': '#F5F9FF', 'border': 1, 'num_format': '#,##,##0',
        'font_size': 9, 'align': 'right'
    })
    f['unit_alt_txt'] = wb.add_format({
        'bg_color': '#F5F9FF', 'border': 1, 'font_size': 9, 'align': 'center'
    })
    f['unit_alt_txt_l'] = wb.add_format({
        'bg_color': '#F5F9FF', 'border': 1, 'font_size': 9, 'align': 'left'
    })
    f['unit_summary'] = wb.add_format({
        'bold': True, 'bg_color': '#EAF0F7', 'border': 1, 'font_size': 10
    })

    # Blank / spacer
    f['blank'] = wb.add_format({'border': 0})
    f['blank_border'] = wb.add_format({'border': 1, 'bg_color': '#FFFFFF'})

    return f


# ── Helper to write a two-column labelled table ───────────────────────────────
def _write_table(ws, start_row, items, f, show_header=False,
                 header_a='Particulars', header_b='Amount (₹)'):
    """items: list of (label, value, val_fmt_key) or (label, value)"""
    r = start_row
    if show_header:
        ws.write(r, 0, header_a, f['col_hdr'])
        ws.write(r, 1, header_b, f['col_hdr'])
        r += 1
    for item in items:
        lbl_text = item[0]
        val_data = item[1]
        fmt_key  = item[2] if len(item) > 2 else 'val_num'
        if lbl_text == '':
            ws.write_blank(r, 0, None, f['blank_border'])
            ws.write_blank(r, 1, None, f['blank_border'])
        else:
            lbl_fmt = f['lbl_indent'] if lbl_text.startswith('  ') or lbl_text.startswith('   ') else f['lbl']
            ws.write(r, 0, lbl_text, lbl_fmt)
            if val_data == '':
                ws.write_blank(r, 1, None, f['blank_border'])
            elif isinstance(val_data, str):
                ws.write(r, 1, val_data, f.get(fmt_key, f['val_txt']))
            else:
                ws.write(r, 1, val_data, f.get(fmt_key, f['val_num']))
        r += 1
    return r


# ─────────────────────────────────────────────────────────────────────────────
# MAIN FUNCTION
# ─────────────────────────────────────────────────────────────────────────────
def generate_output_excel(results):
    """Generate professional audit-ready output Excel. Returns BytesIO."""
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {'in_memory': True})
    f   = _make_formats(wb)
    r   = results
    t   = r['thresholds']
    generated_on = datetime.now().strftime('%d-%b-%Y  %H:%M')

    # =========================================================================
    # SHEET 1 – DASHBOARD
    # =========================================================================
    ws = wb.add_worksheet('📊 Dashboard')
    ws.hide_gridlines(2)
    ws.set_zoom(90)

    ws.set_column('A:A', 2)    # left margin
    ws.set_column('B:B', 22)
    ws.set_column('C:C', 22)
    ws.set_column('D:D', 22)
    ws.set_column('E:E', 22)
    ws.set_column('F:F', 2)    # right margin

    # ── Banner ──
    ws.set_row(0, 8)
    ws.set_row(1, 40)
    ws.set_row(2, 18)
    ws.merge_range('B2:E2',
                   '🏗️  ICAI Revenue Recognition Tool  —  Audit-Ready Output',
                   f['main_title'])
    ws.merge_range('B3:E3',
                   'Project: {}     |     Generated: {}'.format(r['project_name'], generated_on),
                   f['sub_title'])
    ws.set_row(3, 8)

    # ── Section: Project Overview ──
    ws.set_row(4, 22)
    ws.merge_range('B5:E5', '  PROJECT OVERVIEW', f['section_hdr'])

    def kpi_block(ws, row, col, label, value, val_fmt='kpi_val'):
        ws.set_row(row,     20)
        ws.set_row(row + 1, 28)
        ws.write(row,     col, label, f['kpi_lbl'])
        ws.write(row + 1, col, value, f[val_fmt])

    # Row of 4 KPIs  (rows 5-6, cols B-E = 1-4)
    kpi_block(ws, 5, 1, 'Total Saleable Area (sq.ft)', '{:,.0f}'.format(r['total_saleable']))
    kpi_block(ws, 5, 2, 'Total Area Sold (sq.ft)',    '{:,.0f}'.format(r['area_sold']))
    kpi_block(ws, 5, 3, 'Sale Consideration (Agreements)', _inr(r['total_agreement_value']))
    kpi_block(ws, 5, 4, 'Pending Realisation', _inr(r['pending_realisation']),
              'kpi_val_amber' if r['pending_realisation'] > 0 else 'kpi_val')

    ws.set_row(7, 10)

    # ── Section: Revenue Recognition Summary ──
    ws.set_row(8, 22)
    ws.merge_range('B9:E9', '  REVENUE RECOGNITION SUMMARY', f['section_hdr'])

    kpi_block(ws, 9,  1, 'Stage of Completion', _pct(r['stage_of_completion']))
    kpi_block(ws, 9,  2, 'Revenue Recognised',  _inr(r['revenue_to_recognise']),
              'kpi_val_green' if r['revenue_to_recognise'] > 0 else 'kpi_val')
    kpi_block(ws, 9,  3, 'Cost of Revenue (P&L)', _inr(r['cost_of_revenue']))
    profit_fmt = 'kpi_val_green' if r['profit'] >= 0 else 'kpi_val_red'
    kpi_block(ws, 9,  4, 'Profit / (Loss)', _inr(r['profit']), profit_fmt)

    ws.set_row(11, 10)

    # ── Section: Inventory & Disclosures ──
    ws.set_row(12, 22)
    ws.merge_range('B13:E13', '  INVENTORY & DISCLOSURES', f['section_hdr'])

    kpi_block(ws, 13, 1, 'Inventory – Unsold Units', _inr(r['inventory_unsold_units']))
    kpi_block(ws, 13, 2, 'WIP – Sold Units',         _inr(r['wip_sold_units']))
    kpi_block(ws, 13, 3, 'Unbilled Revenue',          _inr(r['unbilled_revenue']),
              'kpi_val_amber' if r['unbilled_revenue'] > 0 else 'kpi_val')
    loss_fmt = 'kpi_val_red' if r['expected_loss'] > 0 else 'kpi_val_green'
    kpi_block(ws, 13, 4, 'Expected Loss [Para 5.7]',
              _inr(r['expected_loss']) if r['expected_loss'] > 0 else 'NIL', loss_fmt)

    ws.set_row(15, 10)

    # ── Section: Threshold Status ──
    ws.set_row(16, 22)
    ws.merge_range('B17:E17', '  THRESHOLD VALIDATION  (Para 5.3)', f['section_hdr'])

    ws.set_row(17, 20)
    ws.set_row(18, 30)
    # Construction
    ws.write(17, 1, 'Construction Completion (≥ 25%)', f['kpi_lbl'])
    ws.write(17, 2, 'Area Sold (≥ 25% of Total Saleable)', f['kpi_lbl'])
    ws.write(17, 3, 'All Thresholds Met ?', f['kpi_lbl'])
    ws.write(17, 4, 'Eligible Contracts', f['kpi_lbl'])
    ws.write(18, 1,
             '✅  {:.1f}%'.format(t['construction_pct']) if t['construction_pass'] else '❌  {:.1f}%'.format(t['construction_pct']),
             f['pass'] if t['construction_pass'] else f['fail'])
    ws.write(18, 2,
             '✅  {:.1f}%'.format(t['area_sold_pct']) if t['area_pass'] else '❌  {:.1f}%'.format(t['area_sold_pct']),
             f['pass'] if t['area_pass'] else f['fail'])
    ws.write(18, 3,
             '✅  YES' if r['all_thresholds_met'] else '❌  NO',
             f['pass'] if r['all_thresholds_met'] else f['fail'])
    eligible_count = int(r['units']['is_eligible'].sum()) if (
        isinstance(r['units'], pd.DataFrame) and 'is_eligible' in r['units'].columns) else 0
    total_count = len(r['units']) if isinstance(r['units'], pd.DataFrame) else 0
    ws.write(18, 4, '{} of {} units'.format(eligible_count, total_count), f['kpi_val'])

    ws.set_row(19, 10)

    # ── Warnings (if any) ──
    if r['warnings']:
        ws.set_row(20, 22)
        ws.merge_range('B21:E21', '  ⚠️  WARNING FLAGS', f['warn_lbl'])
        for idx, w in enumerate(r['warnings']):
            ws.set_row(21 + idx, 18)
            ws.merge_range(21 + idx, 1, 21 + idx, 4,
                           '  ⚠  ' + w, f['warn_lbl'])

    # =========================================================================
    # SHEET 2 – REVENUE WORKING
    # =========================================================================
    ws2 = wb.add_worksheet('Revenue Working')
    ws2.hide_gridlines(2)
    ws2.set_column('A:A', 48)
    ws2.set_column('B:B', 26)
    ws2.set_row(0, 30)
    ws2.merge_range('A1:B1', 'Revenue Recognition Working Sheet', f['sheet_title'])
    ws2.merge_range('A2:B2', 'Project: {}    |    Para 5.1 – 5.6 ICAI Guidance Note (Revised 2012)'.format(r['project_name']), f['sub_title'])
    ws2.set_row(2, 16)

    items = [
        ('Total Estimated Project Cost',          r['est_total'],             'val_num'),
        ('Total Cost Incurred to Date',            r['total_incurred'],        'val_num'),
        ('Stage of Completion (%)',                r['stage_of_completion'],   'val_pct'),
        ('', '', ''),
        ('Construction Completion %',              t['construction_pct'] / 100, 'val_pct'),
        ('Area Sold %',                            t['area_sold_pct'] / 100,  'val_pct'),
        ('Construction Threshold ≥ 25% [Para 5.3(b)]',
         '✅  YES' if t['construction_pass'] else '❌  NO',
         'pass' if t['construction_pass'] else 'fail'),
        ('Area Sold Threshold ≥ 25% [Para 5.3(c)]',
         '✅  YES' if t['area_pass'] else '❌  NO',
         'pass' if t['area_pass'] else 'fail'),
        ('All Para 5.3 Thresholds Met?',
         '✅  YES' if r['all_thresholds_met'] else '❌  NO',
         'pass' if r['all_thresholds_met'] else 'fail'),
        ('', '', ''),
        ('Eligible Revenue (Sum of Eligible Contracts)', r['eligible_revenue'],        'val_num'),
        ('Revenue to Recognise [Stage × Eligible, capped]', r['revenue_to_recognise'], 'val_num'),
    ]
    ws2.write(3, 0, 'Particulars', f['col_hdr'])
    ws2.write(3, 1, 'Value / Amount (₹)', f['col_hdr'])
    row = 4
    for item in items:
        lbl_text, val_data, fmt_key = item
        if lbl_text == '':
            ws2.write_blank(row, 0, None, f['blank_border'])
            ws2.write_blank(row, 1, None, f['blank_border'])
        elif fmt_key in ('pass', 'fail'):
            ws2.write(row, 0, lbl_text, f['lbl'])
            ws2.write(row, 1, val_data, f[fmt_key])
        elif fmt_key == 'val_pct':
            ws2.write(row, 0, lbl_text, f['lbl'])
            ws2.write(row, 1, val_data, f['val_pct'])
        else:
            ws2.write(row, 0, lbl_text, f['lbl'])
            ws2.write(row, 1, val_data, f['val_num'])
        row += 1

    # =========================================================================
    # SHEET 3 – PROFIT SUMMARY
    # =========================================================================
    ws3 = wb.add_worksheet('Profit Summary')
    ws3.hide_gridlines(2)
    ws3.set_column('A:A', 48)
    ws3.set_column('B:B', 26)
    ws3.set_row(0, 30)
    ws3.merge_range('A1:B1', 'Profit / Loss Computation', f['sheet_title'])
    ws3.merge_range('A2:B2', 'Period ending date of computation', f['sub_title'])
    ws3.set_row(2, 14)
    ws3.write(3, 0, 'Particulars', f['col_hdr'])
    ws3.write(3, 1, 'Amount (₹)', f['col_hdr'])
    p_items = [
        ('Revenue Recognised',               r['revenue_to_recognise'], 'val_num'),
        ('Less: Cost of Revenue (P&L)',       r['cost_of_revenue'],      'val_num'),
        ('Profit / (Loss) for the Period',    r['profit'],               'val_num'),
    ]
    row = 4
    for lbl_text, v, _ in p_items:
        ws3.write(row, 0, lbl_text, f['lbl'])
        ws3.write(row, 1, v, f['val_num'])
        row += 1

    ws3.set_row(row, 10)
    row += 1
    ws3.merge_range(row, 0, row, 1, 'Cost Bifurcation', f['section_hdr'])
    row += 1
    ws3.write(row, 0, 'Cost Component', f['col_hdr'])
    ws3.write(row, 1, 'Estimated (₹)', f['col_hdr'])
    row += 1
    for line in [
        ('Land Cost',                t['est_land'],          t['land_incurred']),
        ('Construction Cost',        t['est_construction'],  t['construction_incurred']),
        ('Other / Admin Cost',       t['est_other'],         t['other_incurred']),
        ('Total',                    t['est_total'],         t['total_incurred']),
    ]:
        ws3.write(row, 0, line[0], f['lbl'])
        ws3.write(row, 1, line[1], f['val_num'])
        row += 1

    # =========================================================================
    # SHEET 4 – INVENTORY & WIP
    # =========================================================================
    ws4 = wb.add_worksheet('Inventory & WIP')
    ws4.hide_gridlines(2)
    ws4.set_column('A:A', 55)
    ws4.set_column('B:B', 26)
    ws4.set_row(0, 30)
    ws4.merge_range('A1:B1', 'Closing Inventory / WIP  (Per ICAI Guidance Note Illustration)', f['sheet_title'])
    ws4.merge_range('A2:B2', 'Balance Sheet Date', f['sub_title'])
    ws4.set_row(2, 14)
    ws4.write(3, 0, 'Particulars', f['col_hdr'])
    ws4.write(3, 1, 'Amount (₹)', f['col_hdr'])

    if not r['all_thresholds_met']:
        # PCM conditions NOT met — ALL costs go to Balance Sheet as Inventory/WIP, NOTHING to P&L
        w_items = [
            ('Total Cost Incurred to Date',
             r['total_incurred'], 'val_num'),
            ('Less: Cost of Revenue charged to P&L  [PCM not triggered]',
             'NIL', 'val_txt'),
            ('', '', ''),
            ('Closing Inventory / WIP  (Entire cost — Balance Sheet)',
             r['total_closing_inventory'], 'val_num'),
            ('  [PCM conditions Para 5.3 not yet satisfied — no split applied]',
             '', ''),
            ('', '', ''),
            ('Revenue Recognised',  'NIL', 'val_txt'),
            ('Cost of Revenue (P&L)', 'NIL', 'val_txt'),
            ('Profit / (Loss)',       'NIL', 'val_txt'),
        ]
    else:
        w_items = [
            ('Total Cost Incurred',
             r['total_incurred'], 'val_num'),
            ('Less: Cost of Revenue (P&L)',
             r['cost_of_revenue'], 'val_num'),
            ('', '', ''),
            ('A.  Inventory of Unsold Units',
             r['inventory_unsold_units'], 'val_num'),
            ('     Unsold Area:  {:,.0f} sq.ft'.format(r['unsold_area']), '', ''),
            ('     = (Unsold Area ÷ Total Area)  ×  Total Cost Incurred', '', ''),
            ('', '', ''),
            ('B.  WIP of Sold Units  (cost incurred but not yet recognised)',
             r['wip_sold_units'], 'val_num'),
            ('     = (Sold Area ÷ Total Area)  ×  Cost Incurred  −  Cost of Rev', '', ''),
            ('', '', ''),
            ('Total Closing Inventory / WIP  (A + B)',
             r['total_closing_inventory'], 'val_num'),
            ('Cross-check:  Total Cost Incurred  −  Cost of Revenue',
             r['total_incurred'] - r['cost_of_revenue'], 'val_num'),
        ]

    row = 4
    for item in w_items:
        lbl_text, val_data, _ = item
        if lbl_text == '':
            ws4.write_blank(row, 0, None, f['blank_border'])
            ws4.write_blank(row, 1, None, f['blank_border'])
        elif val_data == '':
            ws4.write(row, 0, lbl_text, f['lbl_indent'])
            ws4.write_blank(row, 1, None, f['blank_border'])
        elif isinstance(val_data, str):
            ws4.write(row, 0, lbl_text, f['lbl'])
            ws4.write(row, 1, val_data, f['val_txt'])
        else:
            ws4.write(row, 0, lbl_text, f['lbl'])
            ws4.write(row, 1, val_data, f['val_num'])
        row += 1

    if r['expected_loss'] > 0:
        ws4.set_row(row, 10); row += 1
        ws4.merge_range(row, 0, row, 1, 'Expected Loss Working  [Para 5.7]', f['section_hdr']); row += 1
        ws4.write(row, 0, 'Total Estimated Project Cost', f['lbl'])
        ws4.write(row, 1, r['est_total'], f['val_num']); row += 1
        ws4.write(row, 0, 'Eligible Revenue', f['lbl'])
        ws4.write(row, 1, r['eligible_revenue'], f['val_num']); row += 1
        ws4.write(row, 0, 'Expected Loss (must be recognised immediately)', f['warn_lbl'])
        ws4.write(row, 1, r['expected_loss'], f['warn_val'])

    # =========================================================================
    # SHEET 5 – JOURNAL ENTRIES  (matching app.py tab exactly)
    # =========================================================================
    ws5 = wb.add_worksheet('Journal Entries')
    ws5.hide_gridlines(2)
    ws5.set_column('A:A', 50)
    ws5.set_column('B:B', 22)
    ws5.set_column('C:C', 22)
    ws5.set_row(0, 30)
    ws5.merge_range('A1:C1', 'Journal Entry Sheet', f['sheet_title'])
    ws5.merge_range('A2:C2',
                    'Project: {}    |    As per ICAI Guidance Note on Real Estate Transactions (Revised 2012)'.format(r['project_name']),
                    f['sub_title'])
    ws5.set_row(2, 14)
    ws5.write(3, 0, 'Account', f['col_hdr'])
    ws5.write(3, 1, 'Debit (₹)', f['col_hdr'])
    ws5.write(3, 2, 'Credit (₹)', f['col_hdr'])

    row = 4
    total_debit  = 0
    total_credit = 0

    def je_section(ws, row, text, f):
        ws5.merge_range(row, 0, row, 2, text, f['je_section'])
        return row + 1

    def je_line_dr(ws5, row, acc, amt, f):
        ws5.write(row, 0, acc, f['je_acc_dr'])
        ws5.write(row, 1, amt, f['je_debit'])
        ws5.write_blank(row, 2, None, f['blank_border'])
        return row + 1, amt

    def je_line_cr(ws5, row, acc, amt, f):
        ws5.write(row, 0, acc, f['je_acc_cr'])
        ws5.write_blank(row, 1, None, f['blank_border'])
        ws5.write(row, 2, amt, f['je_credit'])
        return row + 1, amt

    def je_blank(ws5, row, f):
        ws5.write_blank(row, 0, None, f['blank_border'])
        ws5.write_blank(row, 1, None, f['blank_border'])
        ws5.write_blank(row, 2, None, f['blank_border'])
        return row + 1

    if r['revenue_to_recognise'] > 0:
        # ── Entry 1: Revenue Recognition ──────────────────────────────────────
        row = je_section(ws5, row, '1.  Revenue Recognition  [Para 5.1 / 5.4]', f)

        # Cash/Bank debit = min(amount realised, revenue recognised)
        cash_debit = min(r['total_amount_realised'], r['revenue_to_recognise'])
        if cash_debit > 0:
            row, dr = je_line_dr(ws5, row,
                                 '    Cash / Bank A/c  (Amount Realised — capped at Revenue Recognised)',
                                 cash_debit, f)
            total_debit += dr

        # Unbilled Revenue = Revenue recognised - cash received (if rev > cash)
        unbilled = max(0, r['revenue_to_recognise'] - r['total_amount_realised'])
        if unbilled > 0:
            row, dr = je_line_dr(ws5, row,
                                 '    Unbilled Revenue A/c  (Revenue Recognised > Amount Realised)',
                                 unbilled, f)
            total_debit += dr

        # Credit: Revenue from Operations
        row, cr = je_line_cr(ws5, row,
                              '        To  Revenue from Operations A/c',
                              r['revenue_to_recognise'], f)
        total_credit += cr
        row = je_blank(ws5, row, f)

        # ── Entry 2: Cost Recognition ──────────────────────────────────────────
        row = je_section(ws5, row, '2.  Cost Recognition  [Para 5.6]', f)
        row, dr = je_line_dr(ws5, row,
                              '    Cost of Revenue A/c',
                              r['cost_of_revenue'], f)
        total_debit += dr
        row, cr = je_line_cr(ws5, row,
                              '        To  Work-in-Progress / Inventory A/c',
                              r['cost_of_revenue'], f)
        total_credit += cr
        row = je_blank(ws5, row, f)

        # ── Entry 3: Excess cash → Advances (if amount realised > revenue) ────
        excess_cash = max(0, r['total_amount_realised'] - r['revenue_to_recognise'])
        if excess_cash > 0:
            row = je_section(ws5, row, '3.  Excess Cash Received  [Para 9 — Advances from Customers]', f)
            row, dr = je_line_dr(ws5, row,
                                 '    Cash / Bank A/c  (Collections beyond Revenue Recognised)',
                                 excess_cash, f)
            total_debit += dr
            row, cr = je_line_cr(ws5, row,
                                 '        To  Advances from Customers A/c',
                                 excess_cash, f)
            total_credit += cr
            row = je_blank(ws5, row, f)

    else:
        # ── PCM NOT triggered: all cash treated as Advances from Customers ────
        if r['total_amount_realised'] > 0:
            row = je_section(ws5, row,
                             '1.  Collections from Customers  [PCM not yet triggered — Para 5.3]', f)
            row, dr = je_line_dr(ws5, row,
                                 '    Cash / Bank A/c  (Amount Collected from Customers)',
                                 r['total_amount_realised'], f)
            total_debit += dr
            row, cr = je_line_cr(ws5, row,
                                 '        To  Advances from Customers A/c',
                                 r['total_amount_realised'], f)
            total_credit += cr
            row = je_blank(ws5, row, f)
            ws5.merge_range(row, 0, row, 2,
                            '  Note: Revenue deferred — All costs carried as Inventory / WIP on Balance Sheet.',
                            f['warn_lbl'])
            row += 1
            row = je_blank(ws5, row, f)

    # ── Expected Loss provision (always, if applicable) ───────────────────────
    if r['expected_loss'] > 0:
        entry_num = '3' if r['revenue_to_recognise'] > 0 and r['total_amount_realised'] <= r['revenue_to_recognise'] else \
                    '4' if r['revenue_to_recognise'] > 0 else '2'
        row = je_section(ws5, row,
                         '{}.  Expected Loss Provision  [Para 5.7]'.format(entry_num), f)
        row, dr = je_line_dr(ws5, row,
                              '    Loss on Real Estate Project A/c',
                              r['expected_loss'], f)
        total_debit += dr
        row, cr = je_line_cr(ws5, row,
                              '        To  Provision for Expected Loss A/c',
                              r['expected_loss'], f)
        total_credit += cr
        row = je_blank(ws5, row, f)

    # ── Totals row ─────────────────────────────────────────────────────────
    ws5.write(row, 0, 'TOTAL', f['je_total_lbl'])
    ws5.write(row, 1, total_debit,  f['je_total_val'])
    ws5.write(row, 2, total_credit, f['je_total_val'])

    # ── Note if no entries ──────────────────────────────────────────────────
    if total_debit == 0 and total_credit == 0:
        ws5.merge_range(row, 0, row, 2,
                        'No journal entries — Revenue recognition criteria not yet met and no cash collected.',
                        f['warn_lbl'])


    # =========================================================================
    # SHEET 6 – DISCLOSURE NOTE
    # =========================================================================
    ws6 = wb.add_worksheet('Disclosure Note')
    ws6.hide_gridlines(2)
    ws6.set_column('A:A', 55)
    ws6.set_column('B:B', 26)
    ws6.set_row(0, 30)
    ws6.merge_range('A1:B1', 'Disclosure Note — Para 9, ICAI Guidance Note (Revised 2012)', f['sheet_title'])
    ws6.merge_range('A2:B2', 'Project: {}'.format(r['project_name']), f['sub_title'])
    ws6.set_row(2, 14)
    ws6.write(3, 0, 'Disclosure Item', f['col_hdr'])
    ws6.write(3, 1, 'Amount / Detail', f['col_hdr'])
    disc_items = [
        ('Project Name',                          r['project_name'],               'val_txt'),
        ('Method of Revenue Recognition',         'Percentage of Completion Method (POCM)', 'val_txt'),
        ('Basis — ICAI Guidance Note',            'ICAIGuidance Note on Real Estate Transactions (Revised 2012)', 'val_txt'),
        ('Stage of Completion',                   r['stage_of_completion'],        'val_pct'),
        ('Revenue Recognised',                    r['revenue_to_recognise'],       'val_num'),
        ('Cost of Revenue',                       r['cost_of_revenue'],            'val_num'),
        ('Profit / (Loss)',                       r['profit'],                     'val_num'),
        ('Inventory — Unsold Units',              r['inventory_unsold_units'],     'val_num'),
        ('WIP — Sold Units',                      r['wip_sold_units'],             'val_num'),
        ('Total Closing Inventory / WIP',         r['total_closing_inventory'],    'val_num'),
        ('Unbilled Revenue',                      r['unbilled_revenue'],           'val_num'),
        ('Advances from Customers',               r['advances'],                   'val_num'),
    ]
    if r['expected_loss'] > 0:
        disc_items.append(('Expected Loss Recognised  [Para 5.7]', r['expected_loss'], 'val_num'))
    row = 4
    for lbl_text, val_data, fmt_key in disc_items:
        ws6.write(row, 0, lbl_text, f['lbl'])
        if fmt_key == 'val_pct':
            ws6.write(row, 1, val_data, f['val_pct'])
        elif fmt_key == 'val_txt':
            ws6.write(row, 1, str(val_data), f['val_txt'])
        else:
            ws6.write(row, 1, val_data, f['val_num'])
        row += 1

    # =========================================================================
    # SHEET 7 – UNIT ANALYSIS
    # =========================================================================
    ws7 = wb.add_worksheet('Unit Analysis')
    ws7.hide_gridlines(2)
    ws7.set_column('A:A', 12)   # Unit No
    ws7.set_column('B:B', 18)   # Tower
    ws7.set_column('C:C', 18)   # Area (sq.ft)
    ws7.set_column('D:D', 22)   # Agreement Value
    ws7.set_column('E:E', 22)   # Amount Realised
    ws7.set_column('F:F', 14)   # % Realised
    ws7.set_column('G:G', 12)   # 10% Met
    ws7.set_column('H:H', 10)   # Active
    ws7.set_column('I:I', 10)   # Eligible
    ws7.set_row(0, 30)
    ws7.merge_range('A1:I1', 'Unit-wise Contract Analysis', f['sheet_title'])
    ws7.merge_range('A2:I2',
                    'Project: {}    |    Only units with entered data are shown'.format(r['project_name']),
                    f['sub_title'])
    ws7.set_row(2, 14)
    headers = ['Unit No', 'Tower / Phase', 'Area (Sq.ft)', 'Agreement Value (₹)',
               'Amount Realised (₹)', '% Realised', '10% Met', 'Active', 'Eligible']
    for c, h in enumerate(headers):
        ws7.write(3, c, h, f['unit_hdr'])

    units_df = r['units']
    row = 4
    if isinstance(units_df, pd.DataFrame) and len(units_df) > 0:
        for idx, u in units_df.iterrows():
            alt = idx % 2 == 1
            num_fmt = f['unit_alt'] if alt else f['unit_num']
            txt_fmt = f['unit_alt_txt'] if alt else f['unit_txt']
            txl_fmt = f['unit_alt_txt_l'] if alt else f['unit_txt_l']

            ws7.write(row, 0, str(u.get('unit_no', '')), txl_fmt)
            ws7.write(row, 1, str(u.get('tower', '')), txl_fmt)
            ws7.write(row, 2, float(u.get('saleable_area', 0) or 0), num_fmt)
            ws7.write(row, 3, float(u.get('agreement_value', 0) or 0), num_fmt)
            ws7.write(row, 4, float(u.get('amount_realised', 0) or 0), num_fmt)
            ws7.write(row, 5, float(u.get('pct_realised_calc', 0) or 0), f['unit_pct'])

            for c, key in [(6, 'threshold_10_met'), (7, 'is_active'), (8, 'is_eligible')]:
                flag = bool(u.get(key, False))
                ws7.write(row, c, 'YES' if flag else 'NO',
                          f['unit_yes'] if flag else f['unit_no'])
            row += 1

        # Summary row
        eligible_c = int(units_df['is_eligible'].sum()) if 'is_eligible' in units_df.columns else 0
        active_c   = int(units_df['is_active'].sum())   if 'is_active'   in units_df.columns else 0
        ws7.set_row(row, 10); row += 1
        ws7.merge_range(row, 0, row, 8,
                        'Total Units: {}  |  Active: {}  |  Eligible: {}  |  Non-Eligible: {}'.format(
                            len(units_df), active_c, eligible_c, len(units_df) - eligible_c),
                        f['unit_summary'])
    else:
        ws7.merge_range(row, 0, row, 8, 'No unit data found.', f['warn_lbl'])

    # =========================================================================
    # SHEET 8 – COMPLIANCE REPORT
    # =========================================================================
    ws8 = wb.add_worksheet('Compliance Report')
    ws8.hide_gridlines(2)
    ws8.set_column('A:A', 18)
    ws8.set_column('B:B', 55)
    ws8.set_column('C:C', 16)
    ws8.set_row(0, 30)
    ws8.merge_range('A1:C1', 'Compliance Report — ICAI Guidance Note on Real Estate Transactions (Revised 2012)', f['sheet_title'])
    ws8.set_row(2, 14)
    ws8.write(3, 0, 'Para Reference', f['col_hdr'])
    ws8.write(3, 1, 'Requirement', f['col_hdr'])
    ws8.write(3, 2, 'Status', f['col_hdr'])
    checks = [
        ('Para 5.3(b)', 'Construction & dev. cost incurred >= 25% of total estimated cost',
         t['construction_pass']),
        ('Para 5.3(c)', 'Area secured by agreements >= 25% of total saleable area',
         t['area_pass']),
        ('Para 5.3(d)', 'At least 10% of agreement value realised per active contract',
         t['realisation_pass']),
        ('Para 5.4',    'Revenue recognised using POCM',
         r['all_thresholds_met']),
        ('Para 5.6',    'Cost incurred method used for stage of completion',   True),
        ('Para 5.7',    'Expected loss recognised in full (if applicable)',    True),
        ('Para 9',      'Disclosure note prepared with all required items',    True),
    ]
    for i, (para, req, status) in enumerate(checks):
        ws8.write(i + 4, 0, para, f['lbl'])
        ws8.write(i + 4, 1, req,  f['lbl'])
        ws8.write(i + 4, 2, '✅  COMPLIANT' if status else '❌  NON-COMPLIANT',
                  f['pass'] if status else f['fail'])

    # =========================================================================
    # SHEET 9 – WARNINGS (only if any)
    # =========================================================================
    if r['warnings']:
        ws9 = wb.add_worksheet('⚠ Warnings')
        ws9.hide_gridlines(2)
        ws9.set_column('A:A', 80)
        ws9.set_row(0, 30)
        ws9.merge_range('A1:A1', 'Warning Flags & Audit Alerts', f['sheet_title'])
        for i, w in enumerate(r['warnings']):
            ws9.set_row(i + 1, 20)
            ws9.write(i + 1, 0, '  ⚠  ' + w, f['warn_lbl'])

    wb.close()
    buf.seek(0)
    return buf
