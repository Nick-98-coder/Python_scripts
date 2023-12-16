import os
import pandas as pd
from PIL import Image

def convert_images(input_path, max_size_kb=200):
    output_path = os.path.join(input_path, "webp_folder")

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    file_list = os.listdir(input_path)

    for file_name in file_list:
        if file_name.lower().endswith(('.jpg', '.jpeg')):
            input_file_path = os.path.join(input_path, file_name)
            output_file_path = os.path.join(output_path, os.path.splitext(file_name)[0] + '.webp')

            # Open the image file
            with Image.open(input_file_path) as img:
                # Convert to webp format
                img.save(output_file_path, 'webp', quality=90)

                # Check if the size is within the specified limit
                while os.path.getsize(output_file_path) > max_size_kb * 1024:
                    # If the size exceeds the limit, decrease the quality and save again
                    img.save(output_file_path, 'webp', quality=(img.info.get('quality', 90) - 10))

def main():
    # Replace 'your_excel_file.xlsx' with the actual path to your Excel file
    excel_file_path = r'C:\Users\Tnluser\Documents\Python\jpg2webp\jpg2webp.xlsx'

    # Load the Excel file
    df = pd.read_excel(excel_file_path)

    # Assuming the Excel sheet has a column named 'Path' that contains the location of images
    for index, row in df.iterrows():
        input_path = row['Path']
        convert_images(input_path)

if __name__ == "__main__":
    main()
