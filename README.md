# Tucker Tech Online - XLS to CSV converter
Excel to CSV conv command line .NET 4
Requires .NET 4.0+ and Excel installation.

This file will convert XLS/XLSX files to CSV. Also accepts file inputs so you can mass convert xls -> csv files.
Command line arguments are as follows:

XLSXtoCSV.exe file.xls file.csv
to read a file
XLSXtoCSV.exe -f C:\path\file.txt

The mass input file is comma delimited, so it should look like:

C:\file.xlsx,C:\file.csv
