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

#Reading network path
def net_loc(n):
    for i in range(2, n+1):
        f_path = sheet['C' + str(i)].value
        print("Network path:", f_path)


path = net_loc(n)
print(type(path))