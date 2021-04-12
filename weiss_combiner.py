import os, re, xlrd, openpyxl

path = r"C:\Users\Alex Thompson\Barbaras Development Inc\NP - Documents\NP-Share\Weiss"
filelist = []
container_list = []
filtered_filelist = []

def extract_containers(any_list):
    loc = (os.path.join(path, "Templates\container_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(sheet.nrows):
        any_list.append(sheet.cell_value(i, 0))

    return any_list


def extract_weiss_files(directory, any_list):
    for root, dirs, files in os.walk(directory):
        for file in files:
            any_list.append(os.path.join(root,file))

    return any_list

def filter_files(string, substr):
    filtered_filelist.append(str for str in string if
                            any(sub in str for sub in substr))

    return filtered_filelist

print([i for i in filter_files(extract_weiss_files(path, filelist), extract_containers(container_list))])
