import os, xlrd
from settings import PATH2

def extract_weiss_files():
    filelist = []
    filtered_filelist = []

    loc = (os.path.join(PATH2, "Templates\container_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    container_list = list(zip(sheet.col_values(0, start_rowx=1, end_rowx=None), 
                        sheet.col_values(1, start_rowx=1, end_rowx=None)))
    
    container_list_dict = dict(container_list)
    container_list_dict_values = list(container_list_dict.keys())

    for root, dirs, files in os.walk(PATH2):
        for file in files:
            filelist.append(os.path.join(root,file))

    for str in filelist:
        for sub in container_list_dict_values:
            if sub in str:
                filtered_filelist.append(str)
                # filtered_filelist.append(str for str in string if
                #                         any(sub in str for sub in substr))

    return filtered_filelist, container_list_dict

    
