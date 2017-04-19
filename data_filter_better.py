from openpyxl import load_workbook
from openpyxl import Workbook

def sheet_filter(sheet_ranges, ws):
    mark = False

    for row in sheet_ranges.rows:
        if mark == True and 1 == row[1].value:
            #ws.append([row[2], ])
            row[1].value = 2
            mark = False
        if 0 == row[1].value:
            mark = True
        ws.append(row)
    return ws


wb = load_workbook(filename = 'filter.xlsx')
output = Workbook(write_only=True)


sheet_ranges = wb['sub27_n']
ws = output.create_sheet('sub27_n')
sheet_filter(sheet_ranges, ws)


sheet_ranges = wb['sub27_p']
ws = output.create_sheet('sub27_p')
sheet_filter(sheet_ranges, ws)


sheet_ranges = wb['17-18_n']
ws = output.create_sheet('17-18_n')
sheet_filter(sheet_ranges, ws)


sheet_ranges = wb['17-18_p']
ws = output.create_sheet('17-18_p')
sheet_filter(sheet_ranges, ws)

output.save('filter_output.xlsx')
