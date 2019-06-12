# import xlsxwriter
# import shutil
# from xlutils.copy import copy
import openpyxl
from openpyxl.styles.borders import Border, Side


xfile = openpyxl.load_workbook(
    'general_template_10.xlsx',
)

sheet = xfile['Накладная ТОРГ-12']
sheet.merge_cells('A31:C31')
sheet.merge_cells('D31:S31')

sheet.merge_cells('A32:C32')
sheet.merge_cells('D32:S32')

sheet['A31'].value = 'test'
sheet['A31'].border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

sheet['A32'].value = 'test2'
sheet['A32'].border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

sheet['D31'].value = 'testtesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttest'
sheet['D31'].border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
sheet['D32'].value = 'asdfnasdkfnlakdsflasdkfnkasdfasdfkoj'
sheet['D32'].border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)



xfile.save('text2.xlsx')