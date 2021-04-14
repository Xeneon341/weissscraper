import os
import numpy as np

from extract_files import extract_weiss_files
from openpyxl import load_workbook, Workbook
from settings import PATH

extracted_files = extract_weiss_files()
rows = []
data_rows = []
last_empty_row_list = []

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

for last_row in updated_empty_rows:
    sheet.cell(row = last_row - 1, column=5).value = 'Hey!'

book.save(os.path.join(PATH,"sample.xlsx"))
