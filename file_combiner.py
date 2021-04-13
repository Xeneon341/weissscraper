import os

from extract_files import extract_weiss_files
from openpyxl import load_workbook, Workbook
from settings import PATH

extracted_files = extract_weiss_files()
rows = []

for i in extracted_files:
    wb = load_workbook(filename = i)
    sheet_ranges = wb['Formulas']
    last_empty_row = len(list(sheet_ranges.rows))
    cells = sheet_ranges['A1':'L' + str(last_empty_row)]
    for c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12 in cells:
        rows.append((c1.value, c2.value, c3.value, c4.value, c5.value, c6.value,
                    c7.value, c8.value, c9.value, c10.value, c11.value, c12.value))

tupled_rows = tuple(rows)

book = Workbook()
sheet = book.active

for row in tupled_rows:
    sheet.append(row)

book.save(os.path.join(PATH,"sample.xlsx"))
