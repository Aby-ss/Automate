import requests
from rich import print
from rich.panel import Panel

# Mailgun API endpoint and credentials
MAILGUN_API_KEY = 'your_mailgun_api_key'
MAILGUN_DOMAIN = 'your_mailgun_domain'
MAILGUN_API_URL = f'https://api.mailgun.net/v3/{MAILGUN_DOMAIN}/messages'

# Email content
subject = 'Hello from Python!'
body = 'This is the body of the email.'

# Sender and recipient details
sender_email = 'your_email@example.com'
recipient_emails = ['recipient1@example.com', 'recipient2@example.com']

# Send the email
def send_email(sender, recipients, subject, body):
    request_data = {
        'from': sender,
        'to': recipients,
        'subject': subject,
        'text': body
    }
    response = requests.post(
        MAILGUN_API_URL,
        auth=('api', MAILGUN_API_KEY),
        data=request_data
    )
    if response.status_code == 200:
        success_panel = Panel('[bold green]Email sent successfully![/]', style='green')
        print(success_panel)
    else:
        failure_panel = Panel(f'[bold red]Failed to send email. Response:[/]\n{response.text}', style='red')
        print(failure_panel)

# Call the send_email function
send_email(sender_email, recipient_emails, subject, body)