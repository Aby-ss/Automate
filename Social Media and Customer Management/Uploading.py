import time
from rich.progress import Progress
from instabot import Bot

# Create a bot instance
bot = Bot()

# Login to Instagram
username = "your_username"
password = "your_password"
bot.login(username=username, password=password)

# Path to the image
image_path = "path/to/image.jpg"

with Progress() as progress:
    # Add a task for uploading
    task = progress.add_task("[cyan]Uploading...", total=100)

    # Update the progress and upload the image
    while not progress.finished:
        progress.update(task, advance=1)
        time.sleep(0.1)  # Simulate some processing time

    # Upload the image using the instabot library
    bot.upload_photo(image_path, caption="Your desired caption for the image")

# Logout from Instagram
bot.logout()