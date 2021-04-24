from ctypes.util import find_library
from ctypes import *
import camelot

print(find_library("".join(("gsdll", str(sizeof(c_voidp) * 8), ".dll"))))

tables = camelot.read_pdf('test.PDF')
print(tables)

tables.export('test.csv', f='csv', compress=True) # json, excel, html, sqlite
tables[0]
tables[0].parsing_report
#
tables[0].to_csv('test.csv') # to_json, to_excel, to_html, to_sqlite
# tables[0].df # get a pandas DataFrame!
