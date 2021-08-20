# Weiss Invoices Scraper and Sequence Generator Scripts


<h2>Weiss Multi-Duty File Scraper (file_combiner.py)</h2>

<body>The purpose of this script is to generate a consolidated spreadsheet containing all item allocations
pertaining to each container from a Weiss-Rohlig consolidated surcharge invoice. It first
takes in as input what extract_weiss_files() was able to extract, parses through the desired columns and rows,
and then outputs the consolidated spreadsheet under a given file directory along with a .txt file containing a summary
of invoice numbers it was successfully able to find and scrape.</body>

<h2>Rapidstart Invoice/Item Charge Sequence Generator (rs_prep.py)</h2>

The purpose of this script is to generate two columns of invoice/item charge line coordinates expressed as (inv, item)
in order automate the entering of a consolidated invoice into a Rapidstart import package. The script first starts by reading
from three columns of data inputted in the container_config file by the end-user. The first two columns are all Posted invoice
No.'s and references to all item-applied invoice charge line no.'s, which are expressed in consecutive increments of 10,000 starting from 10,000.
Duplicates of any row from these first two columns represent multiple items being applied from the same invoice charge line. Finally, the third
column represents a filtered list of invoice line no.'s the end-user has already prepared in advanced for any invoice charge line desired for applying
to items.

The script operates first by creating a list beginning with the first item in the "filtered invoice line no.'s to be posted" column.
It then reads down the rows of columns A and B and either repeats the current iterated filter invoice line no. and appends
to the list if there are no changes in sequences from either of the first two columns. Once the break in sequence occurs, the script
increments the current iteration by 10,000 and repeats the process until it has exhausted the number of rows from column's A and B.
