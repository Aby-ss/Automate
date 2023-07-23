from datetime import datetime
import openpyxl
from datetime import datetime
from tabulate import tabulate
from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

from rich import print
from rich import box
from rich.tree import Tree
from rich.text import Text
from rich.align import Align
from rich.panel import Panel
from rich.layout import Layout 
from rich.table import Table

from rich.live import Live
from rich.prompt import Prompt
from rich.progress import track
from rich.progress import BarColumn, Progress, SpinnerColumn, TextColumn

from rich.traceback import install
install(show_locals=True)

layout = Layout()
# Load the Excel file
workbook = openpyxl.load_workbook('Examplery_data.xlsx')

# Select the active sheet
sheet = workbook.active



layout.split_column(
    Layout(name = "Header", size=3),
    Layout(name = "Body"),
    Layout(name = "Footer", size=3)
)

layout["Body"].split_column(
    Layout(name="Data"),
    Layout(name="Charts")
)

layout["Data"].split_column(
    Layout(name="Upper"),
    Layout(name="Lower")
)

layout["Upper"].split_row(
    Layout(name="U1"),
    Layout(name="U2"),
    Layout(name="U3"),
    Layout(name="U4"),
    Layout(name="U5")
)

layout["Lower"].split_row(
    Layout(name="L1"),
    Layout(name="L2"),
    Layout(name="L3"),
    Layout(name="L4"),
    Layout(name="L5")
)

layout["Charts"].split_row(
    Layout(name="Bar Chart"),
    Layout(name="Line Graph")
)

class Header:

    def __rich__(self) -> Panel:
        grid = Table.grid(expand=True)
        grid.add_column(justify="left")
        grid.add_column(justify="center", ratio=1)
        grid.add_column(justify="right")
        grid.add_row(
            "📝", "[b]Automate[/] - [i]Data Dashboard[/]", datetime.now().ctime().replace(":", "[blink]:[/]"),
        )
        return Panel(grid, style="bold white")
    
class Footer:

    def __rich__(self) -> Panel:
        grid = Table.grid(expand=True)
        grid.add_column(justify="center", ratio=1)
        grid.add_row("[i]Empowering Growth through Intelligent Automation[/]")
        return Panel(grid, style="white on black")
    
def gross_values():
    from Excel_file_reader import gross_profit_values_comparison

layout["Header"].update(Header())
layout["Footer"].update(Footer())
layout["U1"].update(gross_values())
print(layout)