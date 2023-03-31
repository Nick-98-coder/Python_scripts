import os
import openpyxl

# load the Excel workbook
wb = openpyxl.load_workbook('C:/Users/Tnluser/Documents/file_rename.xlsx')
sheet = wb.active

# get the number of rows in the sheet
#n = sheet.max_row
n=5
#print("Number of rows:", n)
# iterate through the rows of the sheet use n+1 in (2, 5)

#read old name and new name in excel sheet

def old_n(n):
    for i in range(2, n+1):
        old_name = str(sheet['A' + str(i)].value)
        print(old_name)
        

def new_n(n):
    for i in range(2, n+1):
        new_name = str(sheet['B' + str(i)].value)
        print(new_name)


old_name = old_n(n)

new_name = new_n(n)



