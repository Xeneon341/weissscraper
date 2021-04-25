from ctypes.util import find_library
from ctypes import *
import camelot, os
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

invoice_file = 'DS20210309-1_INV.XLSX'

output_file = os.path.join(cwd, invoice_file)

invoice = pd.read_excel(output_file, sheet_name="INVOICE")

print(invoice)