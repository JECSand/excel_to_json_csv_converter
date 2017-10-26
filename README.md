# excel_to_json_csv_converter

## Overview

A Python script written to make it easy to convert an excel workbook into both csv and json files.
Developed and Tested with Python 2.7 and 3.5 on Debian 9.

## Features
Runs on Linux, Mac and Windows machines.
Script is set up to run on both Python 3.x and 2.x.
The script is set up to automatically install all unmet prerequisite packages using pip.
This script will generated a folder in the output directory for each workbook with csv files for each worksheet.
For json files each worksheet of a workbook will be put into a single json file with a list of it's respective data as the value.
The script can convert both xls and xlsx files.

## How to Use
1. Make sure you have python 2 or 3 installed with pip on your machine
2. git copy the repository
```R
git clone https://github.com/JECSand/excel_to_json_csv_converter.git
```
3. cd into the excel_to_json_csv_converter directory
4. Place the excel file(s) you wish to convert into the input folder
5. Enter the command in the following format to convert a single excel file:
```R
python excel-to-json-csv-converter.py excel_file output_file_type
```
* excel_file is the excel file located in the input folder that you wish to run. Make sure to include the .xls or .xlsx extension.
* output_file_type can either be 'json' or 'csv' and determines the output type. 'csv' is the default value.
6. If the excel file is a string with multiple words, enter in this format:
```R
python excel-to-json-csv-converter.py "excel file.xlsx"
```
7. converted csv/json data will be located in the output folder
8. You can also choose to run the script on all excel files in the input folder at once using all
```R
python3 excel-to-json-csv-converter.py all json
```
