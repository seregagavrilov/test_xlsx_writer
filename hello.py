import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment, Font
from copy import copy
import re

TORG_12_TABLE_CELLS = {
    'row_number': {'A': ['A', 'C']},
    'product_name': {'D': ['D','S']},
    'kod': {'T': ['T', 'W']},
    'measurement': {'X': ['X', 'AB']},
    'okey_kod': {'AC': ['AC', 'AG']},
    'packaging_type': {'AH': ['AH', 'AL']},
    'place': {'AM': ['AM', 'AQ']},
    'count_in_place': {'AR': ['AR', 'AV']},
    'mass': {'AW': ['AW', 'BA']},
    'count': {'BB': ['BB', 'BG']},
    'coast': {'BH': ['BH', 'BP']},
    'sum_without_vat': {'BQ':['BQ', 'BW']},
    'vat': {'BX': ['BX', 'CA']},
    'vat_sum': {'CB': ['CB', 'CH']},
    'sum_with_vat': {'CI': ['CI', 'CQ']}
}

TORG_12_TABLE_RESULT_CELLS = {

}

TORG_12_CELLS = {

}


def fill_cells(sheet):
    sheet['A7'] = 'Поставщикт'
    sheet['I14'] = 'Закупщик'
    sheet['I14'] = 'Поставщикт'
    sheet['I16'] = 'Закупщик'
    sheet['I18'] = 'Основание'
    sheet['AX26'] = 'Номер документа'
    sheet['BI26'] = 'Дата создания тн'
    sheet['K33'] = 7
    # sheet['AR31'] = 'Итого количество'
    # sheet['AR32'] = 'Всего количество'
    # sheet['BQ32'] = 'Сумма без ндс'


def get_sheet(workk_book):
    return workk_book['стр1']


def style_cell(sheet, cell):
    sheet[cell].border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    sheet[cell].alignment = Alignment(
                    horizontal='general',
                    vertical='bottom',
                    text_rotation=0,
                    wrap_text=False,
                    shrink_to_fit=False,
                    indent=0)
    sheet[cell].font = Font(
        name='Calibri',
        size=11,
                 bold=False,
                 italic=False,
                 vertAlign=None,
                 underline='none',
                 strike=False,
                 color='FF000000')


def delete_merged_cell(row):
    count = len(sheet.merged_cells.ranges)
    try:
        for i in range(count):
            if re.findall(r'%s' % str(row), sheet.merged_cells.ranges[i].__str__()):
                end_list = len(sheet.merged_cells.ranges) - 1
                el = sheet.merged_cells.ranges[i]
                sheet.merged_cells.ranges[i] = sheet.merged_cells.ranges[end_list]
                sheet.merged_cells.ranges[end_list] = el
                sheet.merged_cells.ranges.remove(sheet.merged_cells.ranges[end_list])
    except IndexError:
        pass


def fill_profuct_table(sheet):
    ws = work_book.active
    last_row = 0
    for row in range(31, 50):
        ws.insert_rows(row)
        delete_merged_cell(row)
        rd = ws.row_dimensions[row]
        rd.height = 12
        for key, val in TORG_12_TABLE_CELLS.items():
            dict_cells = TORG_12_TABLE_CELLS.get(key)
            for simple_cell, merge_cell in dict_cells.items():
                merg_cell = merge_cell[0]+str(row) + ':' + merge_cell[1]+str(row)
                cell = simple_cell + str(row)
                sheet.merge_cells(merg_cell)
                style_cell(sheet, cell)
                sheet[cell].value = 'val' + str(row)
    last_row = 52

    # sheet.merge_cells('D32:X32')
    # sheet['D32:X32'] = 'Товарная накладная имеет приложение на'

    # sheet.merge_cells('BH58:BU58')
    # sheet.merge_cells('X62:AR62')
    # sheet.merge_cells('CC58:CL58')
    # sheet.merge_cells('BE71:BF71')
    # sheet.merge_cells('BI71:BS71')
    # sheet.merge_cells('CC58:CL58')


def copyRange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
    # Loops through selected Rows
    for i in range(startRow, endRow + 1, 1):
        # Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol, endCol + 1, 1):
            rowSelected.append(sheet.cell(row=i, column=j).value)
        # Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)

    return rangeSelected


def pasteRange(startCol, startRow, endCol, endRow, sheetReceiving, copiedData):
    countRow = 0
    for i in range(startRow, endRow + 1, 1):
        countCol = 0
        for j in range(startCol, endCol + 1, 1):
            sheetReceiving.cell(row=i, column=j).value = copiedData[countRow][countCol]
            countCol += 1
        countRow += 1

if __name__ == '__main__':
    work_book = openpyxl.load_workbook(
        'torg-12.xlsm', guess_types=True
    )

    # wb = Workbook()
    # work_book = wb.active
    sheet = work_book['torg-12_sheet (2)']
    sheet_footer = work_book['torg-12_sheet']
    # sheet.merged_cells.ranges.clear()
    fill_profuct_table(sheet)

    # data = copyRange(1, 1, 95, 27, sheet_footer)
    # pasteRange(1, 51, 95, 76, sheet, data)
    from copy import copy

    for row in sheet_footer.rows:
        for cell in row:
            new_cell = sheet.cell(row=cell.row, column=cell.col_idx,
                                      value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)



    work_book.save('test_home_look.xlsx')
