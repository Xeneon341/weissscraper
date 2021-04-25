from ctypes.util import find_library
from ctypes import *
from docx.api import Document
import camelot, os, docx2txt
import pandas as pd

pd.set_option("display.max_columns", None)
cwd = r"C:\Users\alext\Google Drive\weissscraper\static"
print(cwd)

print(find_library("".join(("gsdll", str(sizeof(c_voidp) * 8), ".dll"))))

# tables = camelot.read_pdf('./static/test2.PDF', pages='all', flavor="stream", cols=5, edge_tol=500)
# print(tables)

# tables.export('test2.csv', f='csv', compress=True) # json, excel, html, sqlite
# print(tables[:])
# [n for n in tables].parsing_report
# #
# tables[0].to_csv('test2.csv') # to_json, to_excel, to_html, to_sqlite
# tables[0].df # get a pandas DataFrame!
# for n in tables:
#     print(n.df)

# invoice_file = 'DS20210309-1_INV.XLSX'

# output_file = os.path.join(cwd, invoice_file)

# invoice = pd.read_excel(output_file, sheet_name="INVOICE")

# print(invoice)


# document = Document('INVOICE--20HR5097.docx')
# print(document)
# table = document.tables[0]

# data = []

# keys = None
# for i, row in enumerate(table.rows):
#     text = (cell.text for cell in row.cells)

#     if i == 0:
#         keys = tuple(text)
#         continue
#     row_data = dict(zip(keys, text))
#     data.append(row_data)
#     print (data)

# df = pd.DataFrame(data)

my_text = docx2txt.process("INVOICE--20HR5097.docx")
print(my_text)