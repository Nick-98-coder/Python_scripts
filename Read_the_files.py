import openpyxl 
import os

# read the Excel file
wb = openpyxl.load_workbook('C:/Users/Tnluser/Documents/File_names.xlsx')
sheet = wb.active


folder_path = str(input("enter your path:\n"))

files = os.listdir(folder_path)
n = len(files)
print(n)

for i, file in zip(range(2, n+1), files):
    filename, ext = os.path.splitext(file)
    present_with_ext = filename + ext
    print("filenames: ", present_with_ext)
    Presentfile_cell = sheet.cell(row=i, column=1)
    Presentfile_cell.value = present_with_ext
    wb.save('C:/Users/Tnluser/Documents/File_names.xlsx')