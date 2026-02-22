"""Generate sample filled template with unsold units AB-101 to AB-104."""
import openpyxl

wb = openpyxl.Workbook()

# Sheet 1: PROJECT_MASTER (no approval fields)
ws1 = wb.active
ws1.title = 'PROJECT_MASTER'
ws1['A1'] = 'PROJECT MASTER DATA'
ws1['A2'] = ''
ws1['A3'] = 'Field'
ws1['B3'] = 'Value'
fields = [
    ('Project Name', 'Sunrise Heights'),
    ('Total Saleable Area (Sq.ft)', 100000),
    ('Estimated Land Cost', 50000000),
    ('Estimated Construction Cost', 120000000),
    ('Estimated Other / Admin Cost', 15000000),
    ('Total Estimated Project Cost', 185000000),
    ('Total Estimated Project Revenue', 350000000),
]
for i, (f, v) in enumerate(fields):
    ws1['A{}'.format(i+4)] = f
    ws1['B{}'.format(i+4)] = v

# Sheet 2: UNIT_WISE_DATA
# AB-101 to AB-104 are UNSOLD (no agreement, no realisation)
# Other units are sold with varying realisations
ws2 = wb.create_sheet('UNIT_WISE_DATA')
ws2['A1'] = 'UNIT WISE SALES DATA'
ws2['A2'] = ''
headers = ['Unit No', 'Tower / Phase', 'Saleable Area (Sq.ft)', 'Agreement Value',
           'Amount Realised Till Date', '% Realised', '10% Threshold Met',
           'Eligible Contract', 'Contract Status']
for i, h in enumerate(headers):
    ws2.cell(row=3, column=i+1, value=h)

# Sold units
units = [
    ('A-201', 'Tower A', 1200, 7200000, 2160000, '', '', '', 'Active'),
    ('A-202', 'Tower A', 1000, 6000000, 1200000, '', '', '', 'Active'),
    ('A-203', 'Tower A', 1500, 9000000, 2700000, '', '', '', 'Active'),
    ('A-204', 'Tower A', 1100, 6600000, 3300000, '', '', '', 'Active'),
    ('B-101', 'Tower B', 1300, 7800000, 2340000, '', '', '', 'Active'),
    ('B-102', 'Tower B', 1400, 8400000, 4200000, '', '', '', 'Active'),
    ('B-103', 'Tower B', 900,  5400000, 1620000, '', '', '', 'Active'),
    ('B-201', 'Tower B', 1100, 6600000, 1980000, '', '', '', 'Active'),
    ('C-101', 'Tower C', 1200, 7200000, 1440000, '', '', '', 'Active'),
    ('C-102', 'Tower C', 800,  4800000, 960000,  '', '', '', 'Active'),
    ('C-103', 'Tower C', 1000, 6000000, 3000000, '', '', '', 'Active'),
    ('C-201', 'Tower C', 1500, 9000000, 4500000, '', '', '', 'Active'),
    ('D-101', 'Tower D', 1200, 7200000, 2160000, '', '', '', 'Active'),
    ('D-102', 'Tower D', 1000, 6000000, 1800000, '', '', '', 'Active'),
    ('D-103', 'Tower D', 1300, 7800000, 3900000, '', '', '', 'Active'),
    ('D-201', 'Tower D', 1100, 6600000, 1980000, '', '', '', 'Active'),
    ('D-202', 'Tower D', 1400, 8400000, 2520000, '', '', '', 'Active'),
    ('E-101', 'Tower E', 1200, 7200000, 2160000, '', '', '', 'Active'),
    ('E-102', 'Tower E', 900,  5400000, 2700000, '', '', '', 'Active'),
    ('E-103', 'Tower E', 1300, 7800000, 2340000, '', '', '', 'Active'),
    ('E-201', 'Tower E', 1100, 6600000, 1980000, '', '', '', 'Active'),
]

for i, u in enumerate(units):
    for j, v in enumerate(u):
        ws2.cell(row=i+4, column=j+1, value=v)

# Unsold units AB-101 to AB-104 (no agreement value, not active sale)
# These units have area but no buyer yet
unsold = [
    ('AB-101', 'Tower AB', 1500, 0, 0, '', '', '', 'Active'),
    ('AB-102', 'Tower AB', 1200, 0, 0, '', '', '', 'Active'),
    ('AB-103', 'Tower AB', 1400, 0, 0, '', '', '', 'Active'),
    ('AB-104', 'Tower AB', 1300, 0, 0, '', '', '', 'Active'),
]
offset = len(units) + 4
for i, u in enumerate(unsold):
    for j, v in enumerate(u):
        ws2.cell(row=offset + i, column=j+1, value=v)

# Sheet 3: COST_INCURRED
ws3 = wb.create_sheet('COST_INCURRED')
ws3['A1'] = 'COST INCURRED TO DATE'
ws3['A2'] = ''
ws3['A3'] = 'Cost Component'
ws3['B3'] = 'Amount'
costs = [
    ('Land Cost Incurred', 50000000),
    ('Construction Cost Incurred', 42000000),
    ('Other / Admin Cost Incurred', 5000000),
    ('Total Cost Incurred', 97000000),
]
for i, (c, v) in enumerate(costs):
    ws3['A{}'.format(i+4)] = c
    ws3['B{}'.format(i+4)] = v

wb.save(r"c:\Users\Jimi Modi\Desktop\AI Class\AICA Level 2\CAPSTONE PROJECT\sample_filled.xlsx")
print("Sample file with unsold units AB-101 to AB-104 created.")
