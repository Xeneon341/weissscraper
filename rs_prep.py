import numpy as np
import csv
import xlrd
import os
from settings import PATH

loc = (os.path.join(PATH, "Templates\container_config.xls"))

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(1)
sheet.cell_value(0, 0)

dict_keys = ["posted_invoice_nos", "posted_invoice_line_nos_applied", "filtered_inv_lines_to_post"]

dict_sequence = {}

zipped_dict_sequence = list(zip(dict_keys, range(len(dict_keys))))

for i, r in zipped_dict_sequence:
    dict_sequence[i] = list(map(lambda x: int(x) if isinstance(x, float) else x, filter(
        lambda x: x != "", list(sheet.col_values(r, start_rowx=1, end_rowx=None)))))


def generate_inv_line_nos_for_items_application(**kwargs):

    posted_invoice_nos = kwargs["posted_invoice_nos"]
    posted_invoice_line_nos_applied = kwargs["posted_invoice_line_nos_applied"]
    filtered_inv_lines_to_post = kwargs["filtered_inv_lines_to_post"]

    indices_of_posted_inv_no_diffs = []
    final_charge_line_assign = []
    posted_invoice_line_nos_applied.append(0)

    zipped_list = list(zip(posted_invoice_nos, posted_invoice_line_nos_applied))

    index_inv_line_diffs = np.insert(np.where(
        np.diff(posted_invoice_line_nos_applied) != 0)[0] + 1, 0, 0)

    for i in range(len(zipped_list)):
        try:
            if (posted_invoice_nos[i] != posted_invoice_nos[i + 1]) and \
                (posted_invoice_line_nos_applied[i] ==
                 posted_invoice_line_nos_applied[i + 1]):
                indices_of_posted_inv_no_diffs.append(i + 1)
        except IndexError:
            pass

    index_diff = np.diff(np.sort(np.append(index_inv_line_diffs,
                                           indices_of_posted_inv_no_diffs)))

    final_output_item_applied_sequence = np.repeat(
        filtered_inv_lines_to_post, index_diff)

    index_inv_line_diffs_charge_assign_diffs = np.insert(np.where(
        np.diff(np.append(final_output_item_applied_sequence, 0)) != 0)[0] + 1, 0, 0)

    for i in np.diff(index_inv_line_diffs_charge_assign_diffs):
        final_charge_line_assign.extend(np.arange(10000, (10000 * i) + 10000, 10000).tolist())

    return final_output_item_applied_sequence, final_charge_line_assign


final_output = generate_inv_line_nos_for_items_application(**dict_sequence)

np.savetxt("out_columns.csv", np.c_[final_output], delimiter=",", fmt="%i")
