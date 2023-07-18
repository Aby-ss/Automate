import pandas as pd
from tabulate import tabulate

# Read the Excel file into a DataFrame
df = pd.read_excel('Examplery_data.xlsx')

# Convert DataFrame to rich table format
table = tabulate(df, headers='keys', tablefmt='fancy_grid')

# Print the rich table
print(table)
