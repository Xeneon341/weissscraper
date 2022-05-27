from ctypes.util import find_library
from ctypes import *
import camelot
import pandas as pd


pd.set_option("display.max_columns", None)
tables = camelot.read_pdf('Summary Templates Report.pdf', flavor='stream', pages='all', cols=4)
tables.export('test.csv', f='csv', compress=True) # json, excel, html, sqlite