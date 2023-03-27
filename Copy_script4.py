import os
import openpyxl 
import shutil

# Open the Excel file
wb = openpyxl.load_workbook('C:/Users/Tnluser/Documents/File_path.xlsx')

# Select the active sheet
sheet = wb.active

#file_finder_function
def search_and_copy(file_name, src_path, dest_path):
    for root, dirs, files in os.walk(src_path):
        for file in files:
            if file == file_name:
                src_file = os.path.join(root, file)
                shutil.copy(src_file, dest_path)
                print(f'{file_name} copied to {dest_path}')
                return
    print(f'{file_name} not found in {src_path}')

# Get the file name and path from the Excel sheet
try:
    for i in range(2,8):
            file_name = sheet['A'+ str(i)].value
            from_path = sheet['B'+ str(i)].value
            to_path = sheet['C'+ str(i)].value

        # Specify the source and destination file paths
            

        # Copy the file from the source to the destination
            search_and_copy(file_name, from_path, to_path)
except TypeError:
    print("Done!")

    



    ## (GPTSceret key) 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