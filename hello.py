import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment, Font
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


def delete_merged_cell(cell, cell2):
    for i in sheet.merged_cells.ranges:
        # if cell in i.__str__():
        if re.findall(r'(?:^|\W)%s(?:$|\W)' % cell, i.__str__()) or \
                re.findall(r'(?:^|\W)%s(?:$|\W)' % cell2, i.__str__()):
            sheet.merged_cells.ranges.remove(i)
            break

def fill_profuct_table(sheet):
    ws = work_book.active
    for row in range(31, 60):
        ws.insert_rows(row)
        rd = ws.row_dimensions[row]
        rd.height = 12
        for key, val in TORG_12_TABLE_CELLS.items():
            dict_cells = TORG_12_TABLE_CELLS.get(key)
            for simple_cell, merge_cell in dict_cells.items():
                merg_cell = merge_cell[0]+str(row) + ':' + merge_cell[1]+str(row)
                cell = simple_cell + str(row)
                # delete_merged_cell(cell, merge_cell[1]+str(row))
                sheet.merge_cells(merg_cell)
                style_cell(sheet, cell)
                sheet[cell].value = 'val' + str(row)


if __name__ == '__main__':
    work_book = openpyxl.load_workbook(
        'torg-12.xlsm',
    )

    # wb = Workbook()
    # work_book = wb.active
    sheet = work_book.active
    # sheet.merged_cells.ranges.clear()
    fill_profuct_table(sheet)
    fill_cells(sheet)
    work_book.save('test_home_look.xlsx')
