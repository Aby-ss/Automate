import requests
from rich import print
from rich.panel import Panel



def send_simple_message():
	return requests.post(
		"https://api.mailgun.net/v3/sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org/messages",
		auth=("api", "39f519e6a04b64c3ad111432e4c9e070-262b213e-b924eb47"),
		data={"from": "Excited User <mailgun@sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org>",
			  "to": ["raoabdulhadi952@gmail.com", "RAO@sandbox767a73cadd3d4ccdbca30676ce7937f7.mailgun.org"],
			  "subject": "Hello",
			  "text": "Testing some Mailgun awesomeness!"})

send_simple_message()
