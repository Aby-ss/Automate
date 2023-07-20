import openpyxl
from tabulate import tabulate

from rich.traceback import install
install(show_locals=True)

# Load the Excel file
workbook = openpyxl.load_workbook('Examplery_data.xlsx')

# Select the active sheet
sheet = workbook.active


# Function to print the cell value of all months
def gross_profit_values():
    # Define the cell numbers for each month
    GROSS_PROFIT_month_cells = {
        'Jan': 'C6',
        'Feb': 'D6',
        'Mar': 'E6',
        'Apr': 'G6',
        'May': 'H6',
        'Jun': 'I6',
        'Jul': 'K6',
        'Aug': 'L6',
        'Sep': 'M6',
        'Oct': 'O6',
        'Nov': 'P6',
        'Dec': 'Q6'
    }
    
    for month, cell_number in GROSS_PROFIT_month_cells.items():
        cell_value = sheet[cell_number].value
        print(f"{month} = {cell_number}")
        print(f"Data: {cell_value}")
        print()  # Empty line for better readability between months

# Call the function to print the cell values of all months
gross_profit_values()