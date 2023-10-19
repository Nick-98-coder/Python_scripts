import os
import subprocess

def html_to_pdf(input_html, output_pdf):
    # Specify the full path to wkhtmltopdf executable
    wkhtmltopdf_path = r'C:\Program Files\poppler\bin\wkhtmltopdf.exe'  # Replace with the actual path on your system

    # Convert HTML to PDF using wkhtmltopdf
    subprocess.run([wkhtmltopdf_path, input_html, output_pdf])

if __name__ == "__main__":
    input_html = input("Enter Html Path:\n")
    output_pdf = input("Enter Output Path:\n")
    html_to_pdf(input_html, output_pdf)
