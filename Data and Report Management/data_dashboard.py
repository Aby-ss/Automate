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
    
# Create an empty database table
database = Table(show_header=True, header_style="bold magenta")
database.add_column("Name")
database.add_column("Age")
database.add_column("Position")
database.add_column("Employee ID")
database.add_column("Email")
database.add_column("Phone")
database.add_column("Address")
database.add_column("Department")
database.add_column("Salary")
database.add_column("Start Date")
database.add_column("Supervisor")
database.add_column("Skills")
database.add_column("Education")

# Function to create a new SQLite database and table
def create_database():
    conn = sqlite3.connect("employee.db")
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS employees (
            name TEXT,
            age INTEGER,
            position TEXT,
            employee_id INTEGER,
            email TEXT,
            phone TEXT,
            address TEXT,
            department TEXT,
            salary REAL,
            start_date TEXT,
            supervisor TEXT,
            skills TEXT,
            education TEXT
        )
        """
    )

    conn.commit()
    conn.close()

# Function to add employee to database
def add_employee_to_database(employee_info):
    conn = sqlite3.connect("employee.db")
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO employees (name, age, position, employee_id, email, phone, address, department,
                               salary, start_date, supervisor, skills, education)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            employee_info["name"],
            employee_info["age"],
            employee_info["position"],
            employee_info["employee_id"],
            employee_info["email"],
            employee_info["phone"],
            employee_info["address"],
            employee_info["department"],
            employee_info["salary"],
            employee_info["start_date"],
            employee_info["supervisor"],
            ", ".join(employee_info["skills"]),
            employee_info["education"]
        )
    )

    conn.commit()
    conn.close()

# Function to retrieve all employees from the database
def retrieve_employees_from_database():
    conn = sqlite3.connect("employee.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM employees")
    rows = cursor.fetchall()

    conn.close()

    return rows

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
        "name": name,
        "age": age,
        "position": position,
        "employee_id": employee_id,
        "email": email,
        "phone": phone,
        "address": address,
        "department": department,
        "salary": salary,
        "start_date": start_date,
        "supervisor": supervisor,
        "skills": skills,
        "education": education
    }

    return employee_info

# Function to display the employee database
def display_database(rows):
    for row in rows:
        database.add_row(row)

    console = Console()
    console.print(Panel.fit("[bold green]Employee Database[/bold green]"))
    console.print(database)

# Main program
def main():
    create_database()

    employee_info = get_employee_info()
    add_employee_to_database(employee_info)

    rows = retrieve_employees_from_database()
    display_database(rows)


layout["Header"].update(Header())
layout["Footer"].update(Footer())

print(layout)
