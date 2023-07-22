from datetime import datetime
from rich import Panel, box

import openpyxl
from tabulate import tabulate
from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

# Load the Excel file
workbook = openpyxl.load_workbook('Examplery_data.xlsx')

# Select the active sheet
sheet = workbook.active

def gross_profit_values(sheet):
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

    def get_previous_month(current_month):
        months_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        current_index = months_list.index(current_month)
        previous_index = (current_index - 1) % len(months_list)
        return months_list[previous_index]

    current_month = datetime.now().strftime("%b")
    previous_month = get_previous_month(current_month)

    panel_content = ""

    for month, cell_number in GROSS_PROFIT_month_cells.items():
        cell_value = sheet[cell_number].value
        panel_content += f"{month} = {cell_value}\n"

    current_value = sheet[GROSS_PROFIT_month_cells[current_month]].value
    previous_value = sheet[GROSS_PROFIT_month_cells[previous_month]].value

    panel_content += f"\nCurrent Month ({current_month}) Value = {current_value}\n"
    panel_content += f"Previous Month ({previous_month}) Value = {previous_value}\n"

    panel_title = "Gross Profit Values"
    panel = Panel.fit(panel_content, title=panel_title, title_align="left", border_style="bold white", box=box.SQUARE)
    print(panel)
    
gross_profit_values(sheet)