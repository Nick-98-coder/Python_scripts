import os
import shutil

permissions = 0o777
source_folder = os.path.abspath("C:\Users\Tnluser\Documents\Artwork list\Drop 7.xlsx")
destination_folder =os.path.abspath("C:\Users\Tnluser\Documents\Custom Office Templates")
#file_name="Drop 7.xlsx"
# fetch all files
#for file_name in os.listdir(source_folder):
    # construct full file path
    #source = source_folder + "\\" + file_name
    
    #destination = destination_folder

    #copy only files
#if os.path.isfile(source_folder):
shutil.copy(os.chmod(source_folder,permissions), os.chmod(destination_folder,permissions))
    #print('copied', file_name)
print(source_folder)
print(destination_folder)