

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

    # Logic Mapping — Row numbers below are 1-based (row 1 = headers, data from row 2)
    # B2=Est Cost, B3=Cost Incurred, B4=Stage%, B5=Constr%, B6=Area%,
    # B7=10%Realised?, B8=Approvals?, B9=POCM Met?, B10=Elig Rev, B11=Rev to Recognise,
    # B12=Area Sold sqft, B13=Cost to P&L, B14=Profit, B15=Closing WIP,
    # B16=Amt Realised, B17=Unbilled, B18=Advances
    logic = [
        ('1.1', 'Total Estimated Project Cost',
         "=Project_Master!B8", "Para 5.2"),
        ('1.2', 'Total Cost Incurred to Date',
         "=Cost_Incurred!B7", "Para 5.2"),
        ('1.3', 'Stage of Completion (%)',
         "=C3/C2", "Para 5.6"),
        ('2.1', 'Construction Completion % [Para 5.3(b)]',
         "=Cost_Incurred!B4/Project_Master!B5", "Para 5.3(b)"),
        ('2.2', 'Area Sold %  [Para 5.3(c)]',
         "=SUMIF(Sales_Data!J:J,\"Active\",Sales_Data!C:C)/Project_Master!B2", "Para 5.3(c)"),
        ('2.3', 'Min 10% Realisation Met? [Para 5.3(d)]',
         "=IF(COUNTIF(Sales_Data!H:H,\"YES\")>0,\"YES\",\"NO\")", "Para 5.3(d)"),
        ('2.4', 'All Thresholds Met? [Para 5.3]',
         "=IF(AND(C5>=0.25,C6>=0.25,C8=\"YES\"),\"YES\",\"NO\")", "Para 5.3"),
        ('3.1', 'Eligible Revenue (Sum of Eligible Contracts)',
         "=SUMIF(Sales_Data!I:I,\"YES\",Sales_Data!D:D)", "Para 5.3(d)"),
        ('3.2', 'Revenue to Recognise',
         "=IF(C9=\"YES\",C4*C10,0)", "Para 5.1"),
        ('4.1', 'Cost of Revenue — P&L  (NIL if thresholds not met)',
         "=IF(C9=\"YES\",(SUMIF(Sales_Data!J:J,\"Active\",Sales_Data!C:C)/Project_Master!B2)*C3,0)",
         "Para 5.6"),
        ('4.2', 'Unsold Area (Sq.ft)',
         "=Project_Master!B2-SUMIF(Sales_Data!J:J,\"Active\",Sales_Data!C:C)",
         "Inventory"),
        ('4.3', 'Inventory of Unsold Units  (Balance Sheet)',
         "=IF(C9=\"YES\",(C13/Project_Master!B2)*C3,0)",
         "Balance Sheet"),
        ('4.4', 'WIP of Sold Units  (Balance Sheet)',
         "=IF(C9=\"YES\",((Project_Master!B2-C13)/Project_Master!B2)*C3-C12,0)",
         "Balance Sheet"),
        ('4.5', 'Total Closing Inventory / WIP',
         "=IF(C9=\"YES\",C14+C15,C3)",
         "Balance Sheet"),
        ('5.1', 'Profit / (Loss)  [NIL if thresholds not met]',
         "=C11-C12", "P&L"),
        ('6.1', 'Amount Realised (Eligible)',
         "=SUMIF(Sales_Data!I:I,\"YES\",Sales_Data!G:G)", "Cash Basis"),
        ('6.2', 'Unbilled Revenue',
         "=MAX(0,C11-C18)", "Disclosure"),
        ('6.3', 'Advances from Customers',
         "=MAX(0,C18-C11)", "Disclosure"),
        ('7.1', 'Expected Loss  [Para 5.7]',
         "=MAX(0,IF(C2>C10,C2-C10,0))", "Para 5.7"),
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
