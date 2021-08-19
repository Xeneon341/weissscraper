# Rapidstart Invoice/Item Charge Sequence Generator

# The purpose of this script is to generate two columns of invoice/item charge line coordinates expressed as (inv, item)
# in order automate the entering of a consolidated invoice into a Rapidstart import package. The script first starts by reading
# from three columns of data inputted in the container_config file by the end-user. The first two columns are all Posted invoice
# No.'s and references to all item-applied invoice charge line no.'s, which are expressed in consecutive increments of 10,000 starting from 10,000.
# Duplicates of any row from these first two columns represent multiple items being applied from the same invoice charge line. Finally, the third
# column represents a filtered list of invoice line no.'s the end-user has already prepared in advanced for any invoice charge line desired for applying
# to items.

# The script operates first by creating a list beginning with the first item in the filtered invoice line no.'s to be posted column.
# It then reads down the rows of columns A and B and either repeats the current iterated filter invoice line no. and appends
# to the list if there are no changes in sequences from either of the first two columns. Once the break in sequence occurs, the script
# increments the current iteration by 10,000 and repeats the process until it has exhausted the number of rows from column's A and B.

import numpy as np
import csv
import xlrd
import os
from settings import TEMPLATE_PATH

loc = (os.path.join(TEMPLATE_PATH, "container_config.xls"))

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(1)
sheet.cell_value(0, 0)

dict_keys = ["posted_invoice_nos", "posted_invoice_line_nos_applied", "filtered_inv_lines_to_post"]

dict_sequence = {}

zipped_dict_sequence = list(zip(dict_keys, range(len(dict_keys))))

# Used to filter out any blank cells and make any conversions from floating point numbers
# to integers.

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


np.savetxt(os.path.join(TEMPLATE_PATH, "rs_prep_output\\generated_rs_import_output.csv"),
           np.c_[final_output], delimiter=",", fmt="%i")
