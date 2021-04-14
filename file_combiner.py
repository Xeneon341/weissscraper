import os
import numpy as np
import easygui

from extract_files import extract_weiss_files
from openpyxl import load_workbook, Workbook
from settings import PATH2

extracted_files, container_dict = extract_weiss_files()
rows = []
data_rows = []
last_empty_row_list = []
container_list_dict_values = list(container_dict.values())

def dialogue_box():
    file_name = easygui.enterbox("What would like you to call your file?")

    return file_name

for i in extracted_files:
    wb_data_only = load_workbook(filename = i, data_only=True)

    sheet_ranges_data_only = wb_data_only['Formulas']

    last_empty_row = len(list(sheet_ranges_data_only.rows))
    last_empty_row_list.append(last_empty_row)

    cells = sheet_ranges_data_only['A1':'L' + str(last_empty_row)]

    for c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12 in cells:
        rows.append((c1.value, c2.value, c3.value, c4.value, c5.value, c6.value,
                    c7.value, c8.value, c9.value, c10.value, c11.value, c12.value))


tupled_rows = tuple(rows)
tupled_data_rows = tuple(data_rows)
updated_empty_rows = np.cumsum(last_empty_row_list)

book = Workbook()
sheet = book.active

for row in tupled_rows:
    sheet.append(row)

for amount, updated_last_row in zip(container_list_dict_values, updated_empty_rows):
    # Update Invoice Amount
    sheet.cell(row = updated_last_row - 4, column = 8).value = amount

for updated_last_row in updated_empty_rows:
    # Update CM3 Multiplier Cell
    sheet.cell(row = updated_last_row - 2, column = 8).value = \
        '=+H' + str(updated_last_row - 3) + '/F' + str(updated_last_row - 8)
    # Update Minus Duty Cell
    sheet.cell(row = updated_last_row - 3, column = 8).value = \
        '=+H' + str(updated_last_row - 4) + '-L' + str(updated_last_row - 7)
    # Update Freight total
    sheet.cell(row = updated_last_row - 3, column = 7).value = \
        '=+H' + str(updated_last_row - 4) + '-L' + str(updated_last_row - 7)
    # Update Quantity Checks
    sheet.cell(row = updated_last_row - 6, column = 5).value = \
        '=+E' + str(updated_last_row - 8) + '-E' + str(updated_last_row - 7)
    # Update CBM/GW Checks
    sheet.cell(row = updated_last_row - 6, column = 6).value = \
        '=+F' + str(updated_last_row - 8) + '-F' + str(updated_last_row - 7)
    # Update Freight Checks
    sheet.cell(row = updated_last_row - 6, column = 7).value = \
        '=+G' + str(updated_last_row - 8) + '-G' + str(updated_last_row - 7)
    # Update Factory Invoice Total Checks
    sheet.cell(row = updated_last_row - 6, column = 8).value = \
        '=+H' + str(updated_last_row - 8) + '-H' + str(updated_last_row - 7)
    # Update Duty + Tariff Checks
    sheet.cell(row = updated_last_row - 6, column = 12).value = \
        '=+L' + str(updated_last_row - 8) + '-L' + str(updated_last_row - 7)
    # Update Freight Amount. Will equal total of the container charges or invoice
    sheet.cell(row = updated_last_row - 7, column = 7).value = \
        '=+H' + str(updated_last_row - 4)
    # Update Duty + Tariff Amount. It's 0 because additional charges don't have d/t tacked on
    sheet.cell(row = updated_last_row - 7, column = 12).value = 0

for updated_last_row, last_row in zip(updated_empty_rows, last_empty_row_list):
    # Update Quantity SUM
    sheet.cell(row = updated_last_row - 8, column = 5).value = \
        '=+SUM(E' + str(3 + (updated_last_row - last_row)) + ':E' + str(1 + (updated_last_row - 11)) + ')'
    # Update CBM SUM
    sheet.cell(row = updated_last_row - 8, column = 6).value = \
        '=+SUM(F' + str(3 + (updated_last_row - last_row)) + ':F' + str(1 + (updated_last_row - 11)) + ')'
    # Update Freight SUM
    sheet.cell(row = updated_last_row - 8, column = 7).value = \
        '=+SUM(G' + str(3 + (updated_last_row - last_row)) + ':G' + str(1 + (updated_last_row - 11)) + ')'
    # Update Invoice SUM
    sheet.cell(row = updated_last_row - 8, column = 8).value = \
        '=+SUM(H' + str(3 + (updated_last_row - last_row)) + ':H' + str(1 + (updated_last_row - 11)) + ')'
    # Update Duty + Tariff SUM
    sheet.cell(row = updated_last_row - 8, column = 12).value = \
        '=+SUM(L' + str(3 + (updated_last_row - last_row)) + ':L' + str(1 + (updated_last_row - 11)) + ')'
    for row in range((3 + (updated_last_row - last_row)), (1 + (updated_last_row - 11))):
        sheet.cell(row = row, column = 7).value = \
            '=F' + str(row) + '*$H$' + str(updated_last_row - 2)
        sheet.cell(row = row, column = 9).value = \
            '=+K' + str(row) + '+J' + str(row)
        sheet.cell(row = row, column = 12).value = \
            '=H' + str(row) + '*((I' + str(row) + \
            '/100)+0.003464+.00125)'

book.save(os.path.join(PATH2, dialogue_box() + '.xlsx'))
