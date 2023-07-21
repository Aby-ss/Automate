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

layout.split_column(
    Layout(name = "Header", size=3),
    Layout(name = "Body"),
    Layout(name = "Footer", size=3)
)

class Header:

    def __rich__(self) -> Panel:
        grid = Table.grid(expand=True)
        grid.add_column(justify="left")
        grid.add_column(justify="center", ratio=1)
        grid.add_column(justify="right")
        grid.add_row(
            "📝", "[b]Automate[/] - [i]HR Dashboard[/]", datetime.now().ctime().replace(":", "[blink]:[/]"),
        )
        return Panel(grid, style="bold white")
    
class Footer:

    def __rich__(self) -> Panel:
        grid = Table.grid(expand=True)
        grid.add_column(justify="center", ratio=1)
        grid.add_row("[i]Empowering Growth through Intelligent Automation[/]")
        return Panel(grid, style="white on black")
    
    

def add_employee_data_to_file(employee_info, file_path):
    with open(file_path, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(employee_info.values())

def print_employee_data_table(file_path):
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        employee_data = list(reader)

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("Name")
    table.add_column("Age")
    table.add_column("Position")
    table.add_column("Employee ID")
    table.add_column("Email")
    table.add_column("Phone")
    table.add_column("Address")
    table.add_column("Department")
    table.add_column("Salary")
    table.add_column("Start Date")
    table.add_column("Supervisor")
    table.add_column("Skills")
    table.add_column("Education")

    for row in employee_data:
        table.add_row(*row)

    console = Console()
    console.print(Panel.fit("[bold green]Employee Information[/bold green]"))
    console.print(table)

def get_employee_info():
    console = Console()

    console.print(Panel.fit("[bold green]Employee Information[/bold green]"))

    name = Prompt.ask("Enter name: ", style="bold white")
    age = Prompt.ask("Enter age: ", style="bold white")
    position = Prompt.ask("Enter position: ", style="bold white")
    employee_id = Prompt.ask("Enter employee ID: ", style="bold white")
    email = Prompt.ask("Enter email address: ", style="bold white")
    phone = Prompt.ask("Enter phone number: ", style="bold white")
    address = Prompt.ask("Enter address: ", style="bold white")
    department = Prompt.ask("Enter department: ", style="bold white")
    salary = Prompt.ask("Enter salary: ", style="bold white")
    start_date = Prompt.ask("Enter start date: ", style="bold white")
    supervisor = Prompt.ask("Enter supervisor: ", style="bold white")
    skills = Prompt.ask("Enter skills (comma-separated): ", style="bold white").split(",")
    education = Prompt.ask("Enter education: ", style="bold white")

    employee_info = {
        "Name": name,
        "Age": age,
        "Position": position,
        "Employee ID": employee_id,
        "Email": email,
        "Phone": phone,
        "Address": address,
        "Department": department,
        "Salary": salary,
        "Start Date": start_date,
        "Supervisor": supervisor,
        "Skills": ", ".join(skills),
        "Education": education
    }

    return employee_info

file_path = "employee_data.csv"

employee_info = get_employee_info()
add_employee_data_to_file(employee_info, file_path)
print_employee_data_table(file_path)

layout["Header"].update(Header())
layout["Footer"].update(Footer())

print(layout)
