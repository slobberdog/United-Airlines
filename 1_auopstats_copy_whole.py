import openpyxl as xl
from openpyxl import Workbook

path1 = 'C:\\Users\\Leona\\Documents\\GitHub\\United Airlines\\auopstats.xlsx'

workbook = Workbook()


template = xl.load_workbook(filename=path1)
temp_sheet = template["summarySNB"]
temp_sheet1 = template["summaryLNB"]

for row in sheet['B8':'V8']:
    for cell in row:
        temp_sheet[cell.coordinate].value = cell.value

for row in sheet1['B8':'V8']:
    for cell in row:
        temp_sheet1[cell.coordinate].value = cell.value

template.save("C:\\Users\\Leona\\Documents\\GitHub\\United Airlines\\auopstats.xlsx")



