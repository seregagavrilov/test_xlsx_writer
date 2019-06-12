import xlsxwriter
import shutil
from xlutils.copy import copy
import openpyxl


xfile = openpyxl.load_workbook(
    'general_template_10.xlsx',
)

sheet = xfile['Накладная ТОРГ-12']
sheet.merge_cells('A31:C31')
sheet.merge_cells('D31:S31')

sheet['A31'].value = 'test'

sheet['D31'].value = 'testtesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttesttest'



xfile.save('text2.xlsx')