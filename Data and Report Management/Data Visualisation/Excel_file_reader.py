import openpyxl
from tabulate import tabulate

from rich.traceback import install
install(show_locals=True)

# Load the Excel file
workbook = openpyxl.load_workbook('Examplery_data.xlsx')

# Select the active sheet
sheet = workbook.active

# Function to get the value of a specific cell based on its cell number
def get_cell_value(cell_number):
    cell = sheet[cell_number]
    return cell.value

# Example usage: retrieve the value of cell 'A1'
cell_number = 'C7'
cell_value = get_cell_value(cell_number)
print(f"The value of cell {cell_number} is: {cell_value}")