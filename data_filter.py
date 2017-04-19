from openpyxl import load_workbook
from openpyxl import Workbook

wb = load_workbook(filename = 'filter_new.xlsx')
output = Workbook(write_only=True)


sheet_ranges = wb['Sheet1']

ws = output.create_sheet('Sheet1')
mark = False

for row in sheet_ranges.rows:
    if mark == True and 1 == row[0].value:
        #ws.append([row[2], ])
        row[0].value = 2
        mark = False
    if 0 == row[0].value:
        mark = True
    ws.append(row)


sheet_ranges = wb['Sheet2']
ws = output.create_sheet('Sheet2')
mark = False

for row in sheet_ranges.rows:
    if mark == True and 1 == row[0].value:
        #ws.append([row[2], ])
        row[0].value = 2
        mark = False
    if 0 == row[0].value:
        mark = True
    ws.append(row)
    
output.save('filter_new_output.xlsx')
