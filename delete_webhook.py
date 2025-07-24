import requests
import os
from dotenv import load_dotenv

# Load the bot token from your .env file
load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")

# Delete the webhook
url = f"https://api.telegram.org/bot{TOKEN}/deleteWebhook"
response = requests.get(url)

# Print the result
print("Delete webhook status:", response.status_code)
print("Response:", response.text)
