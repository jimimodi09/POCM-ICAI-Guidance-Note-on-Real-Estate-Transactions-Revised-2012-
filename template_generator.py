"""
Excel Template Generator for ICAI Revenue Recognition Tool
Generates a downloadable Excel template with 4 sheets and embedded formulas.
"""
import io
import xlsxwriter


def generate_template():
    """Generate Excel template with 4 sheets. Returns BytesIO buffer."""
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {'in_memory': True})

    # --- Formats ---
    hdr = wb.add_format({'bold': True, 'bg_color': '#1B3A5C', 'font_color': '#FFFFFF',
                         'border': 1, 'font_size': 11, 'text_wrap': True, 'valign': 'vcenter'})
    label_fmt = wb.add_format({'bold': True, 'bg_color': '#EAF0F7', 'border': 1,
                               'font_size': 10, 'text_wrap': True})
    input_fmt = wb.add_format({'bg_color': '#FFFDE7', 'border': 1, 'num_format': '#,##,##0',
                               'font_size': 10})
    text_input = wb.add_format({'bg_color': '#FFFDE7', 'border': 1, 'font_size': 10})
    formula_fmt = wb.add_format({'bg_color': '#E8F5E9', 'border': 1, 'num_format': '#,##,##0',
                                 'font_size': 10, 'bold': True, 'font_color': '#1B5E20'})
    pct_fmt = wb.add_format({'bg_color': '#E8F5E9', 'border': 1, 'num_format': '0.00%',
                             'font_size': 10, 'bold': True, 'font_color': '#1B5E20'})
    yn_fmt = wb.add_format({'bg_color': '#E8F5E9', 'border': 1, 'font_size': 10,
                            'bold': True, 'font_color': '#1B5E20', 'align': 'center'})
    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1B3A5C'})
    subtitle_fmt = wb.add_format({'italic': True, 'font_size': 9, 'font_color': '#666666'})
    locked_hdr = wb.add_format({'bold': True, 'bg_color': '#B71C1C', 'font_color': '#FFFFFF',
                                'border': 1, 'font_size': 11})
    locked_label = wb.add_format({'bg_color': '#FFF3E0', 'border': 1, 'font_size': 10})
    locked_val = wb.add_format({'bg_color': '#E3F2FD', 'border': 1, 'num_format': '#,##,##0',
                                'font_size': 10, 'bold': True})
    locked_pct = wb.add_format({'bg_color': '#E3F2FD', 'border': 1, 'num_format': '0.00%',
                                'font_size': 10, 'bold': True})
    locked_yn = wb.add_format({'bg_color': '#E3F2FD', 'border': 1, 'font_size': 10,
                               'bold': True, 'align': 'center'})

    # ========== SHEET 1: PROJECT_MASTER ==========
    ws1 = wb.add_worksheet('PROJECT_MASTER')
    ws1.hide_gridlines(2)
    ws1.set_column('A:A', 40)
    ws1.set_column('B:B', 25)
    ws1.merge_range('A1:B1', 'PROJECT MASTER DATA', title_fmt)
    ws1.merge_range('A2:B2', 'Fill all yellow cells. Green cells are auto-calculated.', subtitle_fmt)

    labels = [
        ('Project Name', text_input, 'Enter Project Name'),
        ('Total Saleable Area (Sq.ft)', input_fmt, 0),
        ('Estimated Land Cost', input_fmt, 0),
        ('Estimated Construction Cost', input_fmt, 0),
        ('Estimated Other / Admin Cost', input_fmt, 0),
        ('Total Estimated Project Cost', formula_fmt, None),  # formula
        ('Total Estimated Project Revenue', input_fmt, 0),
    ]
    ws1.write('A3', 'Field', hdr)
    ws1.write('B3', 'Value', hdr)
    for i, (lbl, fmt, val) in enumerate(labels):
        r = i + 3  # row 3 onwards (0-indexed)
        ws1.write(r, 0, lbl, label_fmt)
        if lbl == 'Total Estimated Project Cost':
            ws1.write_formula(r, 1, '=SUM(B6:B8)', formula_fmt)
        else:
            ws1.write(r, 1, val, fmt)

    # ========== SHEET 2: UNIT_WISE_DATA ==========
    ws2 = wb.add_worksheet('UNIT_WISE_DATA')
    ws2.hide_gridlines(2)
    cols = ['Unit No', 'Tower / Phase', 'Saleable Area (Sq.ft)', 'Agreement Value',
            'Amount Realised Till Date', '% Realised', '10% Threshold Met',
            'Eligible Contract', 'Contract Status']
    widths = [10, 15, 20, 20, 22, 14, 18, 16, 16]
    ws2.merge_range('A1:I1', 'UNIT WISE SALES DATA', title_fmt)
    ws2.merge_range('A2:I2', 'Enter unit details. Columns F–H are auto-calculated. Set Status to Active or Cancelled.', subtitle_fmt)
    for c, (col, w) in enumerate(zip(cols, widths)):
        ws2.set_column(c, c, w)
        ws2.write(2, c, col, hdr)

    # Dropdown for Status
    ws2.data_validation(3, 8, 502, 8, {
        'validate': 'list', 'source': ['Active', 'Cancelled'],
        'input_title': 'Status', 'input_message': 'Select Active or Cancelled'
    })

    for row in range(3, 503):  # rows 4-503 (0-indexed 3-502) — 500 data rows
        r = row + 1  # Excel 1-indexed row
        for c in range(5):
            ws2.write_blank(row, c, None, text_input if c < 2 else input_fmt)
        # % Realised = Amount Realised / Agreement Value
        ws2.write_formula(row, 5, f'=IF(D{r}=0,"",E{r}/D{r})', pct_fmt)
        # 10% Threshold Met
        ws2.write_formula(row, 6, f'=IF(D{r}=0,"",IF(E{r}/D{r}>=0.1,"YES","NO"))', yn_fmt)
        # Eligible Contract
        ws2.write_formula(row, 7,
            f'=IF(D{r}=0,"",IF(AND(G{r}="YES",I{r}="Active"),"YES","NO"))', yn_fmt)
        ws2.write(row, 8, 'Active', text_input)

    # ========== SHEET 3: COST_INCURRED ==========
    ws3 = wb.add_worksheet('COST_INCURRED')
    ws3.hide_gridlines(2)
    ws3.set_column('A:A', 35)
    ws3.set_column('B:B', 25)
    ws3.merge_range('A1:B1', 'COST INCURRED TO DATE', title_fmt)
    ws3.merge_range('A2:B2', 'Enter actual costs incurred. Total is auto-calculated.', subtitle_fmt)
    ws3.write('A3', 'Cost Component', hdr)
    ws3.write('B3', 'Amount (₹)', hdr)

    cost_items = ['Land Cost Incurred', 'Construction Cost Incurred', 'Other / Admin Cost Incurred',
                  'Total Cost Incurred']
    for i, item in enumerate(cost_items):
        r = i + 3
        ws3.write(r, 0, item, label_fmt)
        if item == 'Total Cost Incurred':
            ws3.write_formula(r, 1, '=SUM(B4:B6)', formula_fmt)
        else:
            ws3.write(r, 1, 0, input_fmt)

    # ========== SHEET 4: COMPUTATION_ENGINE ==========
    ws4 = wb.add_worksheet('COMPUTATION_ENGINE')
    ws4.hide_gridlines(2)
    ws4.set_column('A:A', 8)
    ws4.set_column('B:B', 45)
    ws4.set_column('C:C', 25)
    ws4.set_column('D:D', 18)
    ws4.merge_range('A1:D1', 'COMPUTATION ENGINE – DO NOT EDIT', title_fmt)
    ws4.merge_range('A2:D2', 'All values auto-computed from other sheets.', subtitle_fmt)

    for c, col in enumerate(['Step', 'Description', 'Value', 'ICAI Ref']):
        ws4.write(2, c, col, locked_hdr)

    logic = [
        # ── STEP 1: Cost & Stage ──────────────────────────────────────────────
        ('1.1', 'Total Estimated Project Cost',
         '=PROJECT_MASTER!B9', 'Para 5.2', locked_val),
        ('1.2', 'Total Cost Incurred to Date',
         '=COST_INCURRED!B7', 'Para 5.2', locked_val),
        ('1.3', 'Stage of Completion (%)',
         '=IF(C4=0,0,C5/C4)', 'Para 5.6', locked_pct),

        # ── STEP 2: Threshold checks ─────────────────────────────────────────
        ('2.1', 'Construction Completion %  [Para 5.3(b)]',
         '=IF(PROJECT_MASTER!B7=0,0,COST_INCURRED!B5/PROJECT_MASTER!B7)',
         'Para 5.3(b)', locked_pct),
        ('2.2', 'Total Area Sold (Sq.ft)  [Para 5.3(c)]',
         '=SUMPRODUCT((UNIT_WISE_DATA!I4:I503="Active")*(UNIT_WISE_DATA!C4:C503))',
         'Para 5.3(c)', locked_val),
        ('2.3', 'Area Sold %  [Para 5.3(c)]',
         '=IF(PROJECT_MASTER!B5=0,0,C8/PROJECT_MASTER!B5)',
         'Para 5.3(c)', locked_pct),
        # Para 5.3(d): at least one Active contract with >= 10% realisation
        ('2.4', 'All Thresholds Met?  [Para 5.3 — all three conditions]',
         '=IF(AND(C7>=0.25,C9>=0.25,COUNTIFS(UNIT_WISE_DATA!I4:I503,"Active",UNIT_WISE_DATA!G4:G503,"YES")>0),"YES","NO")',
         'Para 5.3', locked_yn),

        # ── STEP 3: Revenue ───────────────────────────────────────────────────
        ('3.1', 'Eligible Revenue (Sum of Eligible Contracts)',
         '=SUMPRODUCT((UNIT_WISE_DATA!H4:H503="YES")*(UNIT_WISE_DATA!D4:D503))',
         'Para 5.3(d)', locked_val),
        ('3.2', 'Revenue to Recognise  [NIL if thresholds not met]',
         '=IF(C10="YES",MIN(C6*C11,C11),0)',
         'Para 5.1', locked_val),

        # ── STEP 4: Cost & Inventory ──────────────────────────────────────────
        # Cost of Revenue: NIL when thresholds not met (entire cost stays on Balance Sheet)
        ('4.1', 'Cost of Revenue (P&L)  [NIL if Para 5.3 not met]',
         '=IF(C10="YES",IF(PROJECT_MASTER!B5=0,0,C6*(C8/PROJECT_MASTER!B5)*C4),0)',
         'Para 5.6', locked_val),
        ('4.2', 'Unsold Area (Sq.ft)',
         '=PROJECT_MASTER!B5-C8',
         'Inventory', locked_val),
        # Inventory of Unsold: NIL when thresholds not met
        ('4.3', 'Inventory of Unsold Units  (Balance Sheet)',
         '=IF(C10="YES",IF(PROJECT_MASTER!B5=0,0,(C14/PROJECT_MASTER!B5)*C5),0)',
         'Balance Sheet', locked_val),
        # WIP of Sold: NIL when thresholds not met
        ('4.4', 'WIP of Sold Units  (Balance Sheet)',
         '=IF(C10="YES",IF(PROJECT_MASTER!B5=0,0,(C8/PROJECT_MASTER!B5)*C5)-C13,0)',
         'Balance Sheet', locked_val),
        # Total Closing Inventory = full cost incurred when thresholds not met
        ('4.5', 'Total Closing Inventory / WIP',
         '=IF(C10="YES",C15+C16,C5)',
         'Balance Sheet', locked_val),

        # ── STEP 5: Profit ────────────────────────────────────────────────────
        # Profit = 0 when thresholds not met (Revenue = 0, Cost of Revenue = 0)
        ('5.1', 'Profit / (Loss)  [NIL if Para 5.3 not met]',
         '=C12-C13',
         'P&L', locked_val),

        # ── STEP 6: Disclosures ───────────────────────────────────────────────
        ('6.1', 'Amount Realised (Eligible)',
         '=SUMPRODUCT((UNIT_WISE_DATA!H4:H503="YES")*(UNIT_WISE_DATA!E4:E503))',
         'Cash Basis', locked_val),
        # Unbilled = Revenue Recognised - Amount Realised (C19 = Amount Realised)
        ('6.2', 'Unbilled Revenue',
         '=MAX(0,C12-C19)',
         'Disclosure', locked_val),
        # Advances = Amount Realised - Revenue Recognised (C19 = Amount Realised)
        ('6.3', 'Advances from Customers',
         '=MAX(0,C19-C12)',
         'Disclosure', locked_val),

        # ── STEP 7: Expected Loss ─────────────────────────────────────────────
        ('7.1', 'Expected Loss  [Para 5.7]',
         '=IF(C4>C11,C4-C11,0)',
         'Para 5.7', locked_val),
    ]

    for i, (step, desc, formula, ref, fmt) in enumerate(logic):
        r = i + 3
        ws4.write(r, 0, step, locked_label)
        ws4.write(r, 1, desc, locked_label)
        ws4.write_formula(r, 2, formula, fmt)
        ws4.write(r, 3, ref, locked_label)

    ws4.protect()

    wb.close()
    buf.seek(0)
    return buf
