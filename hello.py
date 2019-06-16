# import xlsxwriter
# import shutil
# from xlutils.copy import copy
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl.workbook import Workbook

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

TORG_12_CELLS = {

}

# def fill_table(sheet):
#     sheet.merge_cells('A31:C31')
#     sheet.merge_cells('D31:S31')
#     # sheet.merge_cells('A32:C32')
#     # sheet.merge_cells('D32:S32')
#     sheet.merge_cells('T31:W31')
#     sheet.merge_cells('X31:AB31')
#     sheet.merge_cells('AC31:AG31')
#     sheet.merge_cells('AH31:AL31')
#     sheet.merge_cells('AM31:AQ31')
#     sheet.merge_cells('AM31:AQ31')
#     sheet.merge_cells('AR31:AV31')
#     sheet.merge_cells('AW31:BA31')
#     sheet.merge_cells('BB31:BG31')
#     sheet.merge_cells('BH31:BP31')
#     sheet.merge_cells('BQ31:BW31')
#     sheet.merge_cells('BX31:CA31')
#     sheet.merge_cells('CB31:CH31')
#     sheet.merge_cells('CI31:CQ31')
#
#     sheet['A31'].value = '1'
#     sheet['A31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     # sheet['D31'].value = 'Наименование тест'
#     # sheet['D31'].border = Border(
#     #     left=Side(style='thin'),
#     #     right=Side(style='thin'),
#     #     top=Side(style='thin'),
#     #     bottom=Side(style='thin')
#     # )
#
#     sheet['D31'].value = 'Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест'
#     sheet['D31'].border = Border(
#         left=Side(style='medium'),
#         right=Side(style='medium'),
#         top=Side(style='medium'),
#         bottom=Side(style='medium')
#     )
#     sheet['D31'].alignment = Alignment(horizontal="center", vertical="center")
#
#     sheet['T31'].value = 'Код тест'
#     sheet['T31'].border = Border(
#         left=Side(style='medium'),
#         right=Side(style='medium'),
#         top=Side(style='medium'),
#         bottom=Side(style='medium')
#     )
#
#     sheet['X31'].value = 'Код по Окей тест'
#     sheet['X31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BB31'].value = 12
#     sheet['BB31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BQ31'].value = 100.21
#     sheet['BQ31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BX31'].value = 'Ставка'
#     sheet['BX31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BX31'].value = 'Ставка Ндс'
#     sheet['BX31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['CB31'].value = 'Сумма Ндс'
#     sheet['CB31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['CI31'].value = 'Сумма с ндс'
#     sheet['CI31'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     ws = work_book.active
#     ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
#     ws.insert_rows(32)
#
#     sheet.merge_cells('A32:C32')
#     sheet.merge_cells('D32:S32')
#     # sheet.merge_cells('A32:C32')
#     # sheet.merge_cells('D32:S32')
#     sheet.merge_cells('T32:W32')
#     sheet.merge_cells('X32:AB32')
#     sheet.merge_cells('AC32:AG32')
#     sheet.merge_cells('AH32:AL32')
#     sheet.merge_cells('AM32:AQ32')
#     sheet.merge_cells('AM32:AQ32')
#     sheet.merge_cells('AR32:AV32')
#     sheet.merge_cells('AW32:BA32')
#     sheet.merge_cells('BB32:BG32')
#     sheet.merge_cells('BH32:BP32')
#     sheet.merge_cells('BQ32:BW32')
#     sheet.merge_cells('BX32:CA32')
#     sheet.merge_cells('CB32:CH32')
#     sheet.merge_cells('CI32:CQ32')
#
#     sheet['A32'].value = '2'
#     sheet['A32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['D32'].value = 'Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест'
#     sheet['D32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#     sheet['D32'].alignment = Alignment(horizontal="center", vertical="center")
#
#     sheet['T32'].value = 'Код тест'
#     sheet['T32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['X32'].value = 'Код по Окей тест'
#     sheet['X32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     # sheet['BB32'].value = 12.12
#     # sheet['BB32'].border = Border(
#     #     left=Side(style='thin'),
#     #     right=Side(style='thin'),
#     #     top=Side(style='thin'),
#     #     bottom=Side(style='thin')
#     # )
#
#     sheet['BQ32'].value = 100.21
#     sheet['BQ32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BX32'].value = 'Ставка'
#     sheet['BX32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['BX32'].value = 'Ставка Ндс'
#     sheet['BX32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['CB32'].value = 'Сумма Ндс'
#     sheet['CB32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
#
#     sheet['CI32'].value = 'Сумма с ндс'
#     sheet['CI32'].border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')

    #
    # sheet['A33'].value = '3'
    # sheet['A33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    #
    # sheet[
    #     'D33'].value = 'Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест Наименование тест'
    # sheet['D33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    # sheet['D33'].alignment = Alignment(horizontal="center", vertical="center")
    #
    #
    #
    #
    #
    # # sheet['BB32'].value = 12.12
    # # sheet['BB32'].border = Border(
    # #     left=Side(style='thin'),
    # #     right=Side(style='thin'),
    # #     top=Side(style='thin'),
    # #     bottom=Side(style='thin')
    # # )
    #
    # sheet['BQ33'].value = 100.21
    # sheet['BQ33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    #
    # sheet['BX33'].value = 'Ставка'
    # sheet['BX33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    #
    # sheet['BX33'].value = 'Ставка Ндс'
    # sheet['BX33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    #
    # sheet['CB33'].value = 'Сумма Ндс'
    # sheet['CB33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    #
    # sheet['CI33'].value = 'Сумма с ндс'
    # sheet['CI33'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )
    # sheet['A32'].value = 'Чипсы'
    # sheet['A32'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )

    # sheet['D32'].value = 'asdfnasdkfnlakdsflasdkfnkasdfasdfkoj'
    # sheet['D32'].border = Border(
    #     left=Side(style='thin'),
    #     right=Side(style='thin'),
    #     top=Side(style='thin'),
    #     bottom=Side(style='thin')
    # )

def fill_cells(sheet):
    sheet['A7'] = 'Поставщикт'
    sheet['I14'] = 'Закупщик'
    sheet['I14'] = 'Поставщикт'
    sheet['I16'] = 'Закупщик'
    sheet['I18'] = 'Основание'
    sheet['AX26'] = 'Номер документа'
    sheet['BI26'] = 'Дата создания тн'
    sheet['K33'] = 7


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

    # sheet[cell].font


def fill_profuct_table(sheet):
    count = 0
    ws = work_book.active
    for row in range(31, 50):
        ws.insert_rows(row)
        rd = ws.row_dimensions[row]  # get dimension for row 3
        rd.height = 12
        # rd.customFormat =False
        # rd.customHeight = False
        for key, val in TORG_12_TABLE_CELLS.items():
            dict_cells = TORG_12_TABLE_CELLS.get(key)
            for simple_cell, merge_cell in dict_cells.items():
                merg_cell = merge_cell[0]+str(row) + ':' + merge_cell[1]+str(row)
                cell = simple_cell + str(row)
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

    fill_profuct_table(sheet)
    fill_cells(sheet)
    work_book.save('test_home_look.xlsx')
