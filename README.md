#PDF-to-Excel Converter

This Python script extracts key information from multiple PDF files and saves them in an Excel spreadsheet. It uses the tabula, pandas, pathlib, os, camelot, and openpyxl packages to extract and manipulate data.

Purpose
The purpose of this project is to automate the process of manually copying data from multiple PDF files into an Excel spreadsheet. By using this script, users can save time and reduce errors that might occur when manually transferring data.

How it Works
The script takes all the names of the files in the newFolder directory and returns a list of the files. It then iterates through this list of file names and extracts the key information, making a 2D array. It opens the newExcel.xlsx workbook and appends the extracted data to the sheet pyexcel_sheet1. The following values are extracted from each PDF file and added to the Excel document:

LAN
Name
Date of Birth
Date of Referral
Practice Address
Telephone Number

Dependencies
This script relies on the following packages:
tabula
pandas
pathlib
os
camelot
openpyxl

How to Use
Place the PDF files that you want to extract data from in the newFolder directory.
Run the script in a Python environment that has the required packages installed.
The extracted values will be added to the pyexcel_sheet1 sheet in the newExcel.xlsx document, which will be created in the same directory as the script.

Note
The extracted values will be available in the newExcel.xlsx document. If you want to use a different file name or sheet name, you can modify the script accordingly.
