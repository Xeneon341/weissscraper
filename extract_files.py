import os
import xlrd
import re
from settings import PATH
from openpyxl import load_workbook, Workbook


def extract_weiss_files():
    filelist = []
    filtered_filelist = []
    invoice_list = []
    container_values = []

    loc = (os.path.join(PATH, "Templates\container_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    container_list = list(zip(sheet.col_values(0, start_rowx=1, end_rowx=None),
                              sheet.col_values(1, start_rowx=1, end_rowx=None)))

    container_list_dict = dict(container_list)
    container_list_dict_values = list(filter(lambda x: x != "", list(container_list_dict.keys())))

    for root, dirs, files in os.walk(PATH):
        for file in files:
            filelist.append(os.path.join(root, file))

    for str in filelist:
        for sub in container_list_dict_values:
            if (sub in str) and ('Partial Calcs' not in str):
                pattern = '\d+(?=\.xlsx)'
                filtered_filelist.append(str)
                invoice_list.append(re.findall(pattern, str))
                container_list_dict_values.remove(sub)
                container_values.append(container_list_dict[sub])
                # filtered_filelist.append(str for str in string if
                #                         any(sub in str for sub in substr))

    invoice_list = [n for item in invoice_list for n in item]

    return filtered_filelist, container_list_dict, invoice_list, container_values


def extract_from_export_file():
    last_empty_row_list = []
    loc = (os.path.join(PATH, "Templates\importrs_exp.xlsx"))
    wb_data_only = load_workbook(filename=loc, data_only=True)

    sheet_ranges_data_only = wb_data_only['Purch. Inv. Line']

    last_empty_row_list.append(len(list(sheet_ranges_data_only.rows)))
