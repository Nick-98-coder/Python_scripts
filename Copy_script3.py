import openpyxl 
import shutil

# Open the Excel file
wb = openpyxl.load_workbook('C:/Users/Tnluser/Documents/File_path.xlsx')

# Select the active sheet
sheet = wb.active

# Get the file name and path from the Excel sheet
for i in range(2,8):
        file_name = sheet['A'+ str(i)].value
        from_path = sheet['B'+ str(i)].value
        to_path = sheet['C'+ str(i)].value

    # Specify the source and destination file paths
        #src_path = from_path +"\\"+ file_name
        #dst_path = to_path
        #print("From path:", src_path)
        #print("To path:", dst_path)

    # Copy the file from the source to the destination
        #shutil.copy2(src_path, dst_path)
        shutil.copytree(from_path,to_path)
    # Print a confirmation message
        print(f'File {file_name} has been copied from {src_path} to {dst_path}')