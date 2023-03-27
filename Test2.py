import os

folder_path = "C:/Users/Tnluser/Documents/My project/"
# Replace with the path to your folder

# Get a list of all the files in the folder
files = os.listdir(folder_path)

# Print the file names and their extensions
for file in files:
    filename, extension = os.path.splitext(file)
    final = filename + extension
    print(final)
    #print(f"Filename: {filename}, Extension: {extension}")