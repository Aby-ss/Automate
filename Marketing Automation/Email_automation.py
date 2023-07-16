import requests
from rich import print
from rich.panel import Panel

# Mailgun API endpoint and credentials
MAILGUN_API_KEY = '262b213e-b924eb47'
MAILGUN_DOMAIN = 'sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org'
MAILGUN_API_URL = f'https://api.mailgun.net/v3/{MAILGUN_DOMAIN}/messages'

def send_simple_message():
    return requests.post(
        "https://api.mailgun.net/v3/YOUR_DOMAIN_NAME/messages",
        auth=("api", "262b213e-b924eb47"),
        data={"from": "Rao <mailgun@sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org>",
              "to": ["raoabdulhadi952@gmail.com", "rao@sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org"],
              "subject": "Hello",
              "text": "Testing some Mailgun awesomness!"})
    
    if response.status_code == 200:
        success_panel = Panel('[bold green]Email sent successfully![/]', style='green')
        print(success_panel)
    else:
        failure_panel = Panel(f'[bold red]Failed to send email. Response:[/]\n{response.status_code} - {response.content}', style='red')
        print(failure_panel)
    
send_simple_message()