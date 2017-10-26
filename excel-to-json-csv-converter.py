# John Connor Sanders
# Python Excel to CSV/JSON Converter
# 10/09/2017

import pip
import csv
import json
import os
from os import sys


# Package installer
def install(package):
    print(str(package) + ' package for Python not found, pip installing now....')
    pip.main(['install', package])
    print(str(package) + ' package has been successfully installed for Python.\n Continuing Process...')

# Check for xlrd package and install if missing
try:
    import xlrd
except:
    install('xlrd')
    import xlrd

# Script variables and OS Check
cwd = os.getcwd()
os_system = os.name
if os_system == 'nt':
    input_dir = '\\input\\'
    output_dir = '\\output\\'
    csv_dir = 'csv\\'
    json_dir = 'json\\'
    dir_slash = '\\'
else:
    input_dir = '/input/'
    output_dir = '/output/'
    csv_dir = 'csv/'
    json_dir = 'json/'
    dir_slash = '/'


# Function to check user inputs
def check_inputs(script_args):
    if len(script_args) > 3:
        print('Please check your script parameters and try again!')
        print('You entered too many script parameters!')
        sys.exit(1)
    elif len(script_args) < 2:
        print('Please check your script parameters to ensure you are not missing any.')


# Function to check workbook's csv output directory
def ensure_dir(file_path):
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)


# Function to get all excel files or return one based on user'parameter
def get_excel_files(script_param):
    excel_file_list = []
    if script_param == 'all':
        for (dirpath, dirnames, filesnames) in os.walk(str(os.getcwd()) + input_dir):
            excel_file_list.extend(filesnames)
        return excel_file_list
    else:
        excel_file_list.append(script_param)
        return excel_file_list


# Function to build a field value to row position dictionary for fields that need special processing
def get_special_field_names(field_dict, worksheet, rownum):
    special_field_dict = {}
    for key in field_dict.keys():
        if worksheet.row_values(rownum)[field_dict[key]]:
            special_field = worksheet.row_values(rownum)[field_dict[key]]
            dict_ent = {special_field: field_dict[key]}
            special_field_dict.update(dict_ent)
    return special_field_dict


# Function to clean and transform a special value for csv output
def clean_value(val_type, value_dict, workbook):
    cleaned_value_dict = {}
    for key_value in value_dict.keys():
        if isinstance(key_value, float) or isinstance(key_value, int):
            year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(key_value, workbook.datemode)
            if val_type == 'date':
                cleaned_value = "%02d/%02d/%02d" % (month, day, year)
            elif val_type == "time":
                cleaned_value = "%02d:%02d:%02d" % (hour, minute, sec)
            elif val_type == 'datetime':
                cleaned_value = "%02d/%02d/%04d %02d:%02d:%02d" % (month, day, year, hour, minute, sec)
            else:
                cleaned_value = null
            dict_ent = {cleaned_value: value_dict[key_value]}
            cleaned_value_dict.update(dict_ent)
    return cleaned_value_dict


# Function to checks for and returns only special field dictionaries with data
def check_dicts(dicts):
    used_dicts = []
    for diction in dicts:
        if diction:
            used_dicts.append(diction)
    return used_dicts


# Function to update a data row with the new transformed special values
def update_row_values(dicts, worksheet, rownum):
    checked_dicts = check_dicts(dicts)
    if checked_dicts:
        data_row = worksheet.row_values(rownum)
        for checked_dict in checked_dicts:
            for key in checked_dict.keys():
                data_row[checked_dict[key]] = key
        return data_row
    else:
        return 'None'


# Function to build the datetime, date, and time fields lists based on the data's headers
def headers_datetime_process(headers):
    datetime_fields = {}
    date_fields = {}
    time_fields = {}
    for header in headers:
        if 'date' in str(header).lower() and 'time' in str(header).lower():
            dict_ent = {header: headers.index(str(header))}
            datetime_fields.update(dict_ent)
        elif 'date' in str(header).lower() and 'time' not in str(header).lower():
            dict_ent = {header: headers.index(str(header))}
            date_fields.update(dict_ent)
        elif 'time' in str(header).lower() and 'date' not in str(header).lower():
            dict_ent = {header: headers.index(str(header))}
            time_fields.update(dict_ent)
    return [datetime_fields, date_fields, time_fields]


# Function to build and return processed data rows
def get_processed_data_row(workbook, worksheet, rownum, process_data_objs):
    raw_date_dict = get_special_field_names(process_data_objs[1], worksheet, rownum)
    raw_time_dict = get_special_field_names(process_data_objs[2], worksheet, rownum)
    raw_datetime_dict = get_special_field_names(process_data_objs[0], worksheet, rownum)
    date_dict = clean_value('date', raw_date_dict, workbook)
    time_dict = clean_value('time', raw_time_dict, workbook)
    datetime_dict = clean_value('datetime', raw_datetime_dict, workbook)
    updated_row_vals = update_row_values([date_dict, time_dict, datetime_dict], worksheet, rownum)
    return updated_row_vals


# Function to iter through excel work books and write data to new csv file. Python2.x
def excel_csv_conversion_process_py2(workbook, worksheet, out_dir, worksheet_name):
        with open(out_dir + '{}.csv'.format(worksheet_name),
                  'wb') as your_csv_file:
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
            headers = worksheet.row_values(0)
            process_data_objs = headers_datetime_process(headers)
            wr.writerow(headers)
            for rownum in xrange(1, worksheet.nrows):
                updated_row_vals = get_processed_data_row(workbook, worksheet, rownum, process_data_objs)
                if updated_row_vals != 'None':
                    wr.writerow(updated_row_vals)
                else:
                    wr.writerow(worksheet.row_values(rownum))
        your_csv_file.close()


# Function to iter through excel work books and write data to new csv file. Python3.x
def excel_csv_conversion_process_py3(workbook, worksheet, out_dir, worksheet_name):
        with open(out_dir + '{}.csv'.format(worksheet_name),
                  'w', encoding='utf8', newline='') as your_csv_file:
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
            headers = worksheet.row_values(0)
            process_data_objs = headers_datetime_process(headers)
            wr.writerow(headers)
            for rownum in range(1, worksheet.nrows):
                updated_row_vals = get_processed_data_row(workbook, worksheet, rownum, process_data_objs)
                if updated_row_vals != 'None':
                    wr.writerow(updated_row_vals)
                else:
                    wr.writerow(worksheet.row_values(rownum))
        your_csv_file.close()


# Function to run sub process code common to both python 2 and 3
def execute_common_json_code(rownum, workbook, worksheet, process_data_objs, headers):
    updated_row_vals = get_processed_data_row(workbook, worksheet, rownum, process_data_objs)
    i = 0
    sub_obj = {}
    if updated_row_vals != 'None':
        for val in updated_row_vals:
            if val == '':
                val = None
            sub_obj.update({headers[i]: val})
            i += 1
    else:
        for val in worksheet.row_values(rownum):
            if val == '':
                val = None
            sub_obj.update({headers[i]: val})
            i += 1
    return sub_obj


# Function to process cleaned excel data for python 2.x
def json_sub_process_py2(workbook, worksheet, headers, process_data_objs):
    data_list = []
    for rownum in xrange(1, worksheet.nrows):
        data_list.append(execute_common_json_code(rownum, workbook, worksheet, process_data_objs, headers))
    return data_list


# Function to process cleaned excel data for python 3.x
def json_sub_process_py3(workbook, worksheet, headers, process_data_objs):
    data_list = []
    for rownum in range(1, worksheet.nrows):
        data_list.append(execute_common_json_code(rownum, workbook, worksheet, process_data_objs, headers))
    return data_list


# Function to convert excel to JSON data file
def excel_json_conversion_process(workbook, all_worksheets, out_dir, excel_file_name, version):
    data_obj = {}
    for worksheet_name in all_worksheets:
        worksheet = workbook.sheet_by_name(worksheet_name)
        headers = worksheet.row_values(0)
        process_data_objs = headers_datetime_process(headers)
        if version >= 3:
            data_list = json_sub_process_py3(workbook, worksheet, headers, process_data_objs)
        else:
            data_list = json_sub_process_py2(workbook, worksheet, headers, process_data_objs)
        data_obj.update({worksheet_name: data_list})
    with open(out_dir + '{}.json'.format(excel_file_name), 'w') as your_json_file:
        json.dump(data_obj, your_json_file, indent=4)
    your_json_file.close()


# Function to iter through excel workbooks and handle process according to user inputs and python version
def process_handler(workbook, all_worksheets, out_file_dir, excel_file_name, file_type):
    if file_type == 'csv':
        for worksheet_name in all_worksheets:
            worksheet = workbook.sheet_by_name(worksheet_name)
            out_dir = out_file_dir + csv_dir + excel_file_name + dir_slash
            ensure_dir(out_dir)
            if sys.version_info > (3, 0):
                excel_csv_conversion_process_py3(workbook, worksheet, out_dir, worksheet_name)
            else:
                excel_csv_conversion_process_py2(workbook, worksheet, out_dir, worksheet_name)
    elif file_type == 'json':
        out_dir = out_file_dir + json_dir + dir_slash
        ensure_dir(out_dir)
        if sys.version_info > (3, 0):
            excel_json_conversion_process(workbook, all_worksheets, out_dir, excel_file_name, 3)
        else:
            excel_json_conversion_process(workbook, all_worksheets, out_dir, excel_file_name, 2)
    else:
        print('Unrecognized script parameter for file type. Please enter either csv or json!\n Exiting...')
        sys.exit(1)


# Main functions that handles Excel to CSV Transformation Process
def excel_conversion(script_param, file_type='csv'):
    excel_file_list = get_excel_files(script_param)
    for excel_file in excel_file_list:
        workbook = xlrd.open_workbook(cwd + input_dir + excel_file)
        all_worksheets = workbook.sheet_names()
        excel_file_name = excel_file.split('.')[0]
        out_file_dir = cwd + output_dir
        try:
            process_handler(workbook, all_worksheets, out_file_dir, excel_file_name, file_type)
        except Exception as e:
            print(e)
            print('There was an error in the process!')
            sys.exit(1)
    print(file_type.upper() + ' Conversion process is complete!\nNew Files can be found in ' + cwd + output_dir + '!')


if __name__ == "__main__":
    check_inputs(sys.argv)
    if len(sys.argv) == 2:
        excel_conversion(sys.argv[1])
    elif len(sys.argv) == 3:
        excel_conversion(sys.argv[1], sys.argv[2])
