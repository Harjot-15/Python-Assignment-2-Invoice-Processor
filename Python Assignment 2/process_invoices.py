import os
import re
from docx import Document
import pandas as pd

def read_docx(file_path):
    doc = Document(file_path)
    text = " ".join([paragraph.text for paragraph in doc.paragraphs])
    return text

def process_invoices():
    # My directories
    invoices_directory = r'C:\Python Programs\Python Assignment 2\Invoices'
    output_directory = r'C:\Python Programs\Python Assignment 2\Output'

    # For Creating New Directory - create it
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    invoice_data = []

    for filename in os.listdir(invoices_directory):
        if filename.endswith('.docx') and not filename.startswith('~$'):
            # Open the .docx file
            file_path = os.path.join(invoices_directory, filename)
            contents = read_docx(file_path)

            # Extract the necessary data
            invoice_id = re.findall(r'INV\d+', contents)[0]
            product_section = re.search(r'PRODUCTS(.*?)(SUBTOTAL:)', contents, re.DOTALL).group(1)
            product_lines = re.findall(r'\b\w+\s*\w*:\d+', product_section)
            total_products = sum([int(re.search(r'\d+', line).group()) for line in product_lines])
            money_values = re.findall(r'(\d+\.\d+)', contents)
            subtotal = float(money_values[0])
            tax = float(money_values[1])
            total = float(money_values[2])

            # Append the extracted data to the invoice data list
            invoice_data.append([invoice_id, total_products, subtotal, tax, total])

    # Create a pandas DataFrame from the invoice data
    df = pd.DataFrame(invoice_data, columns=["Invoice ID", "Total Products", "Subtotal", "Tax", "Total"])

    # Write the DataFrame to an Excel file in the output directory
    df.to_excel(os.path.join(output_directory, 'invoices.xlsx'), index=False)

if __name__ == "__main__":
    process_invoices()
