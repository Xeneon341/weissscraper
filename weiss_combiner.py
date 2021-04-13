import os, re, xlrd, openpyxl

path = r"C:\Users\Alex Thompson\Barbaras Development Inc\NP - Documents\NP-Share\Weiss"

def extract_containers(any_list):


    return any_list


def extract_weiss_files(directory):
    filelist = []
    container_list = []
    filtered_filelist = []

    loc = (os.path.join(directory, "Templates\container_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(sheet.nrows):
        container_list.append(sheet.cell_value(i, 0))

    for root, dirs, files in os.walk(directory):
        for file in files:
            filelist.append(os.path.join(root,file))

    for str in filelist:
        for sub in container_list:
            if sub in str:
                filtered_filelist.append(str)
                # filtered_filelist.append(str for str in string if
                #                         any(sub in str for sub in substr))

    return filtered_filelist

print(extract_weiss_files(path))
