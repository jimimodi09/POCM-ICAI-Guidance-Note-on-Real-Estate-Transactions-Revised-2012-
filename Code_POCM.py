

def generate_real_estate_pocm_tool(file_name="ICAI_POCM_Tool.xlsx"):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    workbook = writer.book

    # Formatting Styles
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    formula_fmt = workbook.add_format({'bg_color': '#F2F2F2', 'font_color': '#FF0000'})
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    num_fmt = workbook.add_format({'num_format': '#,##0.00'})

    # --- SHEET 1: PROJECT MASTER ---
    master_data = {
        'Metric': [
            'Project Name', 'Total Saleable Area (Sq.ft)', 'Estimated Land Cost',
            'Estimated Development Rights Cost', 'Estimated Construction Cost',
            'Estimated Borrowing Cost', 'Other Allocable Costs', 'Total Estimated Project Cost',
            'Total Estimated Project Revenue', 'Date of Commencement',
            'Critical Approvals Obtained? (Y/N)', 'Change in Land Use? (Y/N)',
            'Environmental Clearance? (Y/N)', 'Title to Land Confirmed? (Y/N)'
        ],
        'Value': ['Project Alpha', 100000, 50000000, 10000000, 100000000, 10000000, 5000000, 0, 300000000, '2024-04-01', 'Y', 'Y', 'Y', 'Y']
    }
    df_master = pd.DataFrame(master_data)
    df_master.to_excel(writer, sheet_name='Project_Master', index=False)
    sheet_master = writer.sheets['Project_Master']
    # Auto-formula for Total Estimated Cost (Sum of rows 3 to 7)
    sheet_master.write_formula('B8', '=SUM(B3:B7)', num_fmt)

    # --- SHEET 2: UNIT WISE SALES DATA ---
    sales_cols = [
        'Unit No', 'Tower', 'Saleable Area', 'Agreement Value', 'Date', 
        '% Realised', 'Amt Realised', '10% Realised?', 'Eligible Contract?', 'Status'
    ]
    df_sales = pd.DataFrame(columns=sales_cols)
    df_sales.to_excel(writer, sheet_name='Sales_Data', index=False)
    sheet_sales = writer.sheets['Sales_Data']
    
    # Adding formulas for row 2 to 100 for user input
    for row in range(1, 101):
        # Amt Realised: Agreement Value * % Realised
        sheet_sales.write_formula(row, 6, f'=D{row+1}*F{row+1}', num_fmt)
        # 10% Realised Check: IF(Amt Realised >= 10% of Agreement Value, "YES", "NO")
        sheet_sales.write_formula(row, 7, f'=IF(G{row+1}>=(0.1*D{row+1}), "YES", "NO")')
        # Eligible Contract (Para 5.3): IF(Status="Active" AND 10% Realised="YES", "YES", "NO")
        sheet_sales.write_formula(row, 8, f'=IF(AND(J{row+1}="Active", H{row+1}="YES"), "YES", "NO")')

    # --- SHEET 3: COST INCURRED ---
    cost_data = {
        'Cost Component': ['Land', 'Dev Rights', 'Construction', 'Borrowing', 'Other', 'Total Incurred'],
        'Amount': [50000000, 10000000, 30000000, 2000000, 1000000, 0]
    }
    df_cost = pd.DataFrame(cost_data)
    df_cost.to_excel(writer, sheet_name='Cost_Incurred', index=False)
    sheet_cost = writer.sheets['Cost_Incurred']
    sheet_cost.write_formula('B7', '=SUM(B2:B6)', num_fmt)

    # --- SHEET 4: AUTOMATIC COMPUTATION (THE ENGINE) ---
    comp_headers = ['Step', 'Description', 'Value', 'Reference']
    df_comp = pd.DataFrame(columns=comp_headers)
    df_comp.to_excel(writer, sheet_name='Computation', index=False)
    sheet_comp = writer.sheets['Computation']

    # Logic Mapping
    logic = [
        ('1.1', 'Total Estimated Project Cost', "=Project_Master!B8", "Para 5.2"),
        ('1.2', 'Total Cost Incurred to Date', "=Cost_Incurred!B7", "Para 5.2"),
        ('1.3', 'Stage of Completion (POCM %)', "=B3/B2", "Para 5.6"),
        ('2.1', 'Construction Cost Threshold (25%)', "=Cost_Incurred!B4/Project_Master!B5", "Para 5.3(b)"),
        ('2.2', 'Area Sold %', "=SUMIF(Sales_Data!J:J, \"Active\", Sales_Data!C:C)/Project_Master!B2", "Para 5.3(c)"),
        ('2.3', 'Legally Enforceable Approvals?', "=IF(COUNTIF(Project_Master!B11:B14, \"N\")>0, \"NO\", \"YES\")", "Para 5.3(a)"),
        ('3.1', 'POCM Recognition Criteria Met?', "=IF(AND(B5>=0.25, B6>=0.25, B7=\"YES\"), \"YES\", \"NO\")", "Para 5.3"),
        ('4.1', 'Total Agreement Value (Eligible)', "=SUMIF(Sales_Data!I:I, \"YES\", Sales_Data!D:D)", "Para 5.3(d)"),
        ('4.2', 'Cumulative Revenue to Recognise', "=IF(B8=\"YES\", B4*B9, 0)", "Para 5.1"),
        ('5.1', 'Area Sold (Sq.ft)', "=SUMIF(Sales_Data!J:J, \"Active\", Sales_Data!C:C)", "Calculation"),
        ('5.2', 'Cost to be Claimed (P&L)', "=(B11/Project_Master!B2)*B3", "Matching Principle"),
        ('6.1', 'Profit for Period', "=B10-B12", "P&L"),
        ('7.1', 'Closing WIP (Inventory)', "=B3-B12", "Para 4.1"),
        ('8.1', 'Amount Realised (Eligible)', "=SUMIF(Sales_Data!I:I, \"YES\", Sales_Data!G:G)", "Cash Basis"),
        ('8.2', 'Unbilled Revenue', "=MAX(0, B10-B14)", "Disclosure"),
        ('8.3', 'Advances from Customers', "=MAX(0, B14-B10)", "Disclosure")
    ]

    for i, (step, desc, form, ref) in enumerate(logic):
        row = i + 1
        sheet_comp.write(row, 0, step)
        sheet_comp.write(row, 1, desc)
        if "=" in form:
            fmt = percent_fmt if "POCM" in desc or "%" in desc else num_fmt
            sheet_comp.write_formula(row, 2, form, fmt)
        sheet_comp.write(row, 3, ref)

    writer.close()
    return file_name

generate_real_estate_pocm_tool()
