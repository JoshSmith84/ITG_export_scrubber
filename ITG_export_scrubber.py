"""
ITG_export_scrubber.py

Author: Josh Smith

Purpose: Stage ITG export data into a single workbook per client with
separate worksheets based on input data csv's. Format cells based on
width of widest string, freeze header row/format as a table,
and zip for sending. Also edit cell data to remove html,
 and leave only readable data.
 """

# imports
import os
import shutil
import csv
from zipfile import ZipFile
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %'
                                                '(levelname)s - %'
                                                '(message)s'
                    )


# Get input directory (tkinter)

# Iterate through every zip file (outer loop)

# Unzip input
working_dir = "U:\\Joshua\\Work-Stuff\\ITG stuff\\example_exports\\"
input_zip = f"{working_dir}Wyco_ITGRaw.zip"
temp_dir = f"{working_dir}temp\\"
with ZipFile(input_zip, 'r') as zip:
    zip.extractall(temp_dir)


# Delete folders
d_folders = ['attachments', 'documents']
for folder in d_folders:
    if os.path.exists(f"{temp_dir}{folder}"):
        shutil.rmtree(f"{temp_dir}{folder}")


# Delete unneeded csv's

# Create the output Excel workbook file based on name of input zip

# (function; inner loop)
# Iterate through every remaining csv, pulling all data into list of lists
#   while keeping track of headers.
# Go through and delete unneeded columns, any rows that are archived
# Go through those lists inspecting every string and scrubbing html data
#   while leaving true data.
# Output what remains to Sheet in workbook based on name of input csv
# Format sheet

# Zip the output workbook; restart outer loop