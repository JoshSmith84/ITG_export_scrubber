from ITG_Export_Scrubber.itg_scrub_pkg.Application import Application
from ITG_Export_Scrubber.itg_scrub_pkg.MainPage import MainPage
import os
from os.path import basename
import shutil
import csv
import unicodedata
from bs4 import BeautifulSoup
from zipfile36 import ZipFile
from openpyxl import load_workbook
import datetime
import pandas as pd
import xlsxwriter
import traceback


def process_exports(input_zip, post_task, zip_task) -> int:
    """Main processing method. Take a TPG ITG export, unzip it,
    ignore unneeded data, clean left over csv's of empty columns,
    and make it readable

    :param input_zip: any zip file. Non ITG exports will be unzipped,
                    and ignored once contents are detected as invalid.
    :param post_task: Option to either delete or keep input zip
                    when processing is complete.
    :param zip_task: Option to either zip the output or not
                    when processing is complete.
    :return: Integer 0 or 1. 0 means no error occurred.
    1 means error log is present.
    """

    # List of CSV's to extract. Ignore all else
    keep_csv = [
        'applications-licensing.csv', 'backup.csv',
        'backups-managed.csv',
        'battery-backup-ups.csv', 'configurations.csv',
        'domain-hosting.csv',
        'email.csv', 'file-sharing.csv', 'internet-wan.csv',
        'lan.csv', 'passwords.csv', 'printing.csv',
        'vendors.csv', 'voice-pbx-fax.csv', 'wireless.csv',
    ]

    # Column names that if left in first column after processing,
    # should be sorted by in the final sheet
    sorters = ['Name', 'name', 'Hostname', 'Printer Name',
               'Description', 'Vendor Name'
               ]

    # Prep the processing paths
    explode_path = input_zip.split('/')
    new_explode_path = [i.replace('/', '\\') for i in explode_path]
    input_zip = '\\'.join(new_explode_path)
    new_explode_path.pop(-1)
    working_dir = '\\'.join(new_explode_path) + '\\'
    export_dir = working_dir + 'itg_unzipped\\'
    error_log = working_dir + (f'ITG_scrubber_errors_'
                               f'{datetime.date.today()}.txt')

    # Dictionary to contain all customer sheets:dataframe
    sheets_dict = {}

    # Unzip input
    try:
        with ZipFile(input_zip, 'r') as in_zip:
            for file in in_zip.infolist():
                if file.filename in keep_csv:
                    in_zip.extract(file, export_dir)
    except FileNotFoundError:
        MainPage.log_error(error_log, f'{input_zip} not found. '
                                      f'More Info: {traceback.format_exc()}'
                                      f'\n\n'
                           )
        return 1
    except PermissionError:
        MainPage.log_error(error_log, f'{input_zip} '
                                      f'permission denied.'
                                      f' Try Running again as admin. '
                                      f'More Info: {traceback.format_exc()}'
                                      f'\n\n'
                           )
        return 1
    except in_zip.BadZipFile:
        MainPage.log_error(error_log, f'{input_zip}'
                                      f' may be corrupt. '
                                      f'More Info: {traceback.format_exc()}'
                                      f'\n\n'
                           )
        return 1
    except OSError:
        MainPage.log_error(error_log, f'{input_zip} '
                                      f'caused an OS error. '
                                      f' Drive may be full or '
                                      f'path is no longer valid. '
                                      f'More Info: {traceback.format_exc()}'
                                      f'\n\n'
                           )
        return 1

    # Check if the zip is a valid export(export_dir will not exist)
    if os.path.exists(export_dir):
        # Gather list of files
        raw_input_files = [f for f in os.listdir(export_dir)]
    else:
        MainPage.status.set(f'{input_zip} is not a valid ITG export')
        Application.update(Application)
        return 0

    # If Backup-managed is present, ignore backup
    if 'backups-managed.csv' in raw_input_files:
        input_files = [f for f in raw_input_files if f != 'backup.csv']
    else:
        input_files = [f for f in raw_input_files]

    # From any of the csvs, pull customer name from column B
    with open(export_dir + input_files[0], 'r',
              encoding='utf-8') as csv_file:
        headers = csv_file.readline().strip('\n').split(',')
        reader = csv.reader(csv_file)
        customer_name = list(reader)[0][1]

    # Iterate through every remaining csv,
    # and make changes in memory
    for file in input_files:
        MainPage.status.set(f'Processing {customer_name} ...')
        Application.update(Application)
        # Reset columns to delete list
        # (columns that always are deleted/ignored)
        delete_columns = ['id', 'organization', 'Category',
                          'Business Impact',
                          'Client Subject Matter Expert',
                          'Importance', 'archived',
                          'Backup Estimated Start Date',
                          'FlexAssset Review Date',
                          'FlexAsset Review Date',
                          'Backup Radar Reporting Schedule',
                          'hostname', 'manufacturer',
                          'position', 'contact', 'location',
                          'configuration_interfaces',
                          'DHCP Exclusions', 'one_time_password',
                          'Printer Management Login',
                          'installed_by',
                          'Equipment make & Model',
                          'resource_type', 'resource_id',
                          'configuration_status', 'asset_tag',
                          'DHCP Server', 'DHCP Scope',
                          'DHCP Reservations', 'DNS Server(s)',
                          'Default Gateway Device', 'Firewall',
                          'Access Point(s)',
                          'Wireless Controller (Application)',
                          'Wireless Controller (Hardware)',
                          'Management Credentials', 'VLAN #',
                          'Backup Radar Report Recipients (Email)'
                          ' or Link Contacts',
                          'Backup Radar Reporting Notes',
                          'Backup Server/NAS Management Login',
                          'Local Backup Encryption Key',
                          'Backup Copy Job Name',
                          'Backup Copy Target',
                          'Backup Copy Encryption',
                          'Configuration Backup to Cloud Connect?',
                          'SMB Login',
                          ]

        # Continue with unpacking current csv to list of lists
        working_rows = []
        with open(export_dir + file, 'r', encoding='utf-8') as csv_file:
            headers = csv_file.readline().strip('\n').split(',')
            reader = csv.reader(csv_file)
            for row in reader:
                new_row = []
                for cell in row:
                    # Detect html in cell and convert if so.
                    if bool(
                            BeautifulSoup(cell, 'lxml'
                                          ).find()
                    ):
                        cell = BeautifulSoup(cell, 'lxml').text

                    # normalize text
                    cell = unicodedata.normalize(
                        'NFKD', cell
                    )
                    new_row.append(cell)
                working_rows.append(new_row)

        # Find the archive column and keep track of it
        # (it's usually last but not always)
        # Also find config status column
        archive_index = -1
        configuration_status_index = 0
        for index, value in enumerate(headers):
            if value == 'archived':
                archive_index = index
            if file == 'configurations.csv':
                if value == 'configuration_status':
                    configuration_status_index = index

        # go through every row and delete any row with archive set to 'Yes'
        # and any configuration status in configurations csv
        # other than Active
        top_index_current = len(working_rows) - 1
        for index, value in enumerate(reversed(working_rows)):
            if value[archive_index] == 'Yes':
                del working_rows[top_index_current - index]
            if configuration_status_index != 0:
                if value[configuration_status_index] != 'Active':
                    del working_rows[top_index_current - index]

        # Find empty columns
        for i in range(len(headers)):
            i_empty = 0
            for value in working_rows:
                if value[i] == '':
                    i_empty += 1
            if i_empty == len(working_rows):
                delete_columns.append(headers[i])

        # Ignore all blank columns and columns to be "deleted"
        clean_rows = []
        for row in working_rows:
            new_row = []
            for i in range(len(row)):
                if headers[i] not in delete_columns:
                    new_row.append(row[i])
            clean_rows.append(new_row)

        # Clean up headers to match
        new_headers = [h for h in headers if h not in delete_columns]

        # Convert the list of lists to a pandas DataFrame
        columns_dict = {}
        for row in clean_rows:
            for i, value in enumerate(row):
                columns_dict.setdefault(new_headers[i], []).append(value)
        df = pd.DataFrame(columns_dict)

        # Sort any sheets if name column is present and
        # Populate the Pandas dataframe into a dictionary to
        # populate the workbook after all data is processed
        if len(new_headers) > 0:
            if new_headers[0] in sorters:
                sheets_dict.update({file.split('.')[0]: df.sort_values(
                    by=new_headers[0])}
                )
            else:
                sheets_dict.update({file.split('.')[0]: df})
        else:
            continue

    # Populate Workbook with all sheets/tables and data
    wb_file = working_dir + f'{customer_name}_export.xlsx'
    if os.path.exists(wb_file):
        try:
            os.remove(wb_file)
        except PermissionError:
            MainPage.log_error(error_log, f'Attempted '
                                          f'deleting old {wb_file}, '
                                          f' but permission denied.'
                                          f' Try running again as admin. '
                                          f'More Info: {traceback.format_exc()}'
                                          f'\n\n'
                               )
            shutil.rmtree(export_dir)
            return 1
    wb = xlsxwriter.Workbook(wb_file)
    for key, value in sheets_dict.items():
        sheet = wb.add_worksheet(key)
        sheet.add_table(0, 0, value.shape[0], value.shape[1] - 1, {
            'data': value.values.tolist(),
            'columns': [{'header': col} for col in value.columns]
        })
    wb.close()

    # Delete the unzipped original export
    shutil.rmtree(export_dir)

    # Open workbook again with openpyxl and make final adjustments
    # Find and set uniform column width
    wb = load_workbook(wb_file)
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            if max_length > 50:
                max_length = 50
            adjusted_width = (max_length + 2) * 1.2
            sheet.column_dimensions[column].width = adjusted_width

    wb.save(wb_file)

    # Delete or keep unzipped export
    if post_task == 'Delete':
        try:
            os.remove(input_zip)
        except PermissionError:
            MainPage.log_error(error_log, f'Attempted '
                                          f'deleting {input_zip}, '
                                          f' but permission denied.'
                                          f' Try running again as admin. '
                                          f'More Info: {traceback.format_exc()}'
                                          f'\n\n'
                               )
            return 1

    MainPage.status.set(f'Processing of {customer_name} complete.')
    Application.update(Application)

    # To zip or not to zip output file (needs to be zipped for email)
    if zip_task == 'Yes':
        out_zip = f'{working_dir}\\{customer_name}_export.zip'
        if os.path.exists(out_zip):
            try:
                os.remove(out_zip)
            except PermissionError:
                MainPage.log_error(error_log, f'Attempted '
                                              f'deleting {input_zip}, '
                                              f' but permission denied.'
                                              f' Try running again as admin. '
                                              f'More Info: '
                                              f'{traceback.format_exc()}'
                                              f'\n\n'
                                   )
                return 1
        with ZipFile(out_zip, 'w') as f:
            f.write(wb_file, basename(wb_file))
        try:
            os.remove(wb_file)
        except PermissionError:
            MainPage.log_error(error_log, f'Attempted '
                                          f'deleting {wb_file} '
                                          f'as it has been zipped, '
                                          f' but permission denied.'
                                          f' Try running again as admin. '
                                          f'More Info: {traceback.format_exc()}'
                                          f'\n\n'
                               )
    return 0
