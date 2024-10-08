# Import the pdfplumber library, which is used to extract data from PDF files
import pdfplumber as pdfp

# Import the pandas library, which is used to manipulate and analyze data
import pandas as pd

# Define a function called pdf_to_excel that takes two arguments: pdf_file (path to the PDF) and excel_file (path to save Excel file)
def pdf_to_excel(pdf_file, excel_file):
    # Open the PDF file using pdfplumber's open method
    with pdfp.open(pdf_file) as pdf:
        # Initialize an empty list to store all the tables extracted from the PDF pages
        all_tables = []
        
        # Iterate over each page in the PDF
        for page in pdf.pages:
            # Extract tables from the current page using pdfplumber's extract_tables method
            tables = page.extract_tables()
            
            # Iterate over each extracted table on the page
            for table in tables:
                # Check if the table is not empty
                if table:
                    # Convert the table (list of lists) into a pandas DataFrame for better manipulation
                    df = pd.DataFrame(table)
                    
                    # Add the DataFrame to the list of all tables
                    all_tables.append(df)
                    
        # If no tables were found in the entire PDF, create a DataFrame with a message indicating this
        if not all_tables:
            all_tables.append(pd.DataFrame([["No tables were found"]]))
            
        # Create an ExcelWriter object to write the DataFrames to an Excel file using the 'openpyxl' engine
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Iterate over each DataFrame in the list of all tables
            for idx, df in enumerate(all_tables):
                # Write the DataFrame to a separate sheet in the Excel file (naming each sheet as Sheet 1, Sheet 2, etc.)
                df.to_excel(writer, sheet_name=f"Sheet {idx+1}", index=False)
                
# Ask the user for the path to the PDF file they want to convert to Excel
pdf_file_path = input("Please enter the path to your PDF file: ")

# Ask the user for the path to save the Excel file (must include .xlsx extension in the path)
excel_file_path = input("Please enter the path to where you want to save your Excel file (including the .xlsx extension): ")

# Call the pdf_to_excel function with the paths provided by the user to perform the conversion
pdf_to_excel(pdf_file_path, excel_file_path)
