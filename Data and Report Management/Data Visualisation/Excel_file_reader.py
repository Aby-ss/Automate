import pandas as pd

# Define the path to your Excel file
excel_file_path = 'path/to/your/excel/file.xlsx'

# Load the Excel file into a pandas DataFrame
data_frame = pd.read_excel(excel_file_path)

# Extract data from specific sheet
sheet_name = 'Sheet1'  # Replace with your sheet name
sheet_data = data_frame.parse(sheet_name)

# Extract specific columns from the sheet
columns_to_extract = ['Column1', 'Column2', 'Column3']  # Replace with your column names
extracted_data = sheet_data[columns_to_extract]

# Print the extracted data
print(extracted_data)