"""

    Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    Report bugs to @ nikhil.ss4795@gmail.com


    Description :

        Joins multiple csv or xlsx files into a single csv or xlsx files, if
        it contains same kind of data under same headers.

        * Works fine with blank cases also.
        * Makes a note at start which says from which input sheet the data had came.
        * All the input sheets should be of one type only (CSV or XLSX).
        * Works fine with large number of input sheets containing large data.
        

    possibilities :

        1. Multiple XLSX to single XLSX
        2. Multiple XLSX to single CSV
        3. Multiple CSV to single CSV
        4. Multiple CSV to single XLSX

    Input sheet constraints :
        1. Must be either CSV or XLSX (Recommended csv because it is quick)
        2. All the input sheets must have same length and same order of header.
        

    Script goes wrong in cases like :
        1. Input sheets contains both xlsx and csv files.
        2. Input sheets doesn't contains same header or doesn't follow same order in all sheets.
        3. Wrong Input sheets type mentioned.
        

    How to Run :
        1. Keep the input sheets and the script file in one folder.
        2. Run the script file.
        3. Choose the input sheets type.
        4. Choose the output sheet type required.

    Requirements :
        1. python 2 or python 3.
        2. Modules required in python :
            csv, xlrd, xlsxwriter, os, sys, time.
            (To install modules : pip install module_name)
        
"""

import csv
import xlrd, xlsxwriter
import os
import sys
import time

def removeNonAscii(s):
    return "".join(filter(lambda x: ord(x)<128, s))

print("Input sheet types acceptable : ")
print("     1. CSV")
print("     2. XLSX")

try :
    file_type_input = input("\nChoose your input sheet type (1 for CSV) or (2 for XLSX)  : ")
except:
    print("\nError : Wrong input sheet type format choosen.")
    print("        Please input either 1 or 2")
    sys.exit()

if file_type_input == 1 or file_type_input == 2 :
    pass
else:
    print("\nError : Wrong input sheet type format choosen.")
    print("        Please input either 1 or 2")
    sys.exit()


print("\n\nOutput sheet types acceptable : ")
print("     1. CSV")
print("     2. XLSX")


try :
    file_type_output = input("\nChoose your output sheet type (1 for CSV) or (2 for XLSX)  : ")

except:
    print("\nError : Wrong output sheet type format choosen.")
    print("        Please input either 1 or 2")
    sys.exit()

if file_type_output == 1 or file_type_output == 2 :
    pass
else:
    print("\nError : Wrong output sheet type format choosen.")
    print("        Please input either 1 or 2")
    sys.exit()

    

start_time = time.time()

file_names = os.listdir('.')

print("\nScript file and Sheets considered : ")
print(file_names)

if file_type_output == 1 :
    dataout = open('Combined_data_sheet.csv','wb')
    datawrite = csv.writer(dataout)

if file_type_output == 2 :
    output_sheet = xlsxwriter.Workbook('Combined_data_sheet.xlsx')
    worksheet = output_sheet.add_worksheet()

header = []
header.append("Data came from ")
head_count = 0
row_count = 0

if file_type_input == 1:
    for i in range(len(file_names)):
        if file_names[i][-3:] == 'csv':
            if head_count == 0:
                head_count = head_count + 1
                datasheet_for_header = csv.reader(open(file_names[i],'rU'))
                for index,each in enumerate(datasheet_for_header):
                    if index == 0:
                        for data in range(0,len(each)):
                            header.append(each[data])
                    else:
                        break
                if file_type_output == 1 :
                    datawrite.writerow(header)
                if file_type_output == 2 :
                    for data in range(0,len(header)):
                        worksheet.write(row_count, data, header[data])
                    row_count = row_count + 1
                    
            datain = csv.reader(open(file_names[i],'rU'))
            for index,each in enumerate(datain):
                if index == 0:
                    pass
                else:
                    row_data = []
                    for sheet_data in range(0,len(each)):
                        row_data.append(each[sheet_data])
                    row_data = [file_names[i]] + row_data

                    if file_type_output == 1 :
                        datawrite.writerow(row_data)
                    if file_type_output == 2 :
                        for data in range(0,len(row_data)):
                            worksheet.write(row_count, data, row_data[data])
                        row_count = row_count + 1

if file_type_input==2 :
    for i in xrange(len(file_names)):
        if file_names[i][-4:] == 'xlsx':
            if head_count == 0:
                head_count = head_count + 1
                datasheet_for_header = xlrd.open_workbook(file_names[i]).sheet_by_index(0)
                for row in range(0,datasheet_for_header.nrows):
                    if row == 0:
                        for col in range(0,datasheet_for_header.ncols):
                            header.append(datasheet_for_header.cell_value(row,col))
                    else:
                        break
                if file_type_output == 1 :
                    datawrite.writerow(header)
                if file_type_output == 2 :
                    for data in range(0,len(header)):
                        worksheet.write(row_count, data, header[data])
                    row_count = row_count + 1

            xl_reader = xlrd.open_workbook(file_names[i]).sheet_by_index(0)
            for row in range(1,xl_reader.nrows):
                row_data = []
                for j in range(xl_reader.ncols):
                    try:
                        row_data.append(removeNonAscii(xl_reader.cell_value(row, j)))
                    except:
                        row_data.append(xl_reader.cell_value(row, j))
                row_data = [file_names[i]] + row_data

                if file_type_output == 1 :
                    datawrite.writerow(row_data)
                if file_type_output == 2 :
                    for data in range(0,len(row_data)):
                        worksheet.write(row_count, data, row_data[data])
                    row_count = row_count + 1


if file_type_output == 1 :
    dataout.close()

if file_type_output == 2 :
    output_sheet.close()

print("\nTask Completed")
print("     Took around : %s seconds " % (time.time() - start_time))

"""

    Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    Report bugs to @ nikhil.ss4795@gmail.com

"""
