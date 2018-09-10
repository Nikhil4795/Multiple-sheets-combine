# Multiple-sheets-combine
Joins multiple csv or xlsx files into a single csv or xlsx file.


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
