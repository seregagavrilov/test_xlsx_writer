from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active

data = [
    ['Apples', 10000, 5000, 8000, 6000, 1,2,3,4,5,6,4,2,3,4,5],
    ['Apples', 10000, 5000, 8000, 6000, 1,2,3,4,5,6,4,2,3,4,5],
    ['Apples', 10000, 5000, 8000, 6000, 1,2,3,4,5,6,4,2,3,4,5],
    ['Apples', 10000, 5000, 8000, 6000, 1,2,3,4,5,6,4,2,3,4,5],
    ['Apples', 10000, 5000, 8000, 6000, 1,2,3,4,5,6,4,2,3,4,5],
]

# add column headings. NB. these must be strings
ws.append(["Fruit", "2011", "2012", "2013", "2014"])
for row in data:
    ws.append(row)

tab = Table(displayName="Table1", ref="A31:94C")
ws.add_table(tab)
wb.save("table.xlsx")