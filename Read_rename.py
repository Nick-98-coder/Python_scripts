import openpyxl
import os

# Read the Excel file
wb = openpyxl.load_workbook(r'C:\Users\nikhil.srinivas\Documents\filename.xlsx')
sheet = wb.active

# Get the data from the Excel sheet
data = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    data.append(row)

# Iterate through the data and rename files
for row in data:
    filename = row[0]
    folder_path = row[1]
    new_filename = row[2]

    file_path = os.path.join(folder_path, filename)
    file_extension = os.path.splitext(filename)[1]
    new_file_path = os.path.join(folder_path, new_filename + file_extension)

    if os.path.exists(file_path) and os.path.isfile(file_path):
        os.rename(file_path, new_file_path)

# Save the updated Excel file
wb.save(r'C:\Users\nikhil.srinivas\Documents\filename.xlsx')
