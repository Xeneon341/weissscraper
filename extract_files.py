import os, xlrd
from settings import PATH

def extract_weiss_files():
    filelist = []
    container_list = []
    filtered_filelist = []

    loc = (os.path.join(PATH, "Templates\container_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(sheet.nrows):
        container_list.append(sheet.cell_value(i, 0))

    for root, dirs, files in os.walk(PATH):
        for file in files:
            filelist.append(os.path.join(root,file))

    for str in filelist:
        for sub in container_list:
            if sub in str:
                filtered_filelist.append(str)
                # filtered_filelist.append(str for str in string if
                #                         any(sub in str for sub in substr))

    return filtered_filelist
    
