import telebot
from dotenv import load_dotenv
import os

load_dotenv()
TOKEN = os.getenv("BOT_TOKEN")

bot = telebot.TeleBot(TOKEN)

@bot.message_handler(commands=["getid"])
def send_chat_id(message):
    bot.reply_to(message, f"Chat ID: `{message.chat.id}`", parse_mode="Markdown")

print("Bot is running. Send /getid inside the group or channel.")
bot.polling()

'''
# Debug handler
# Print and send telegram group ID when user send a message in the group where bot is added
@bot.message_handler(func=lambda m: True)
def debug_chat_id(message):
    print(f"[DEBUG] Chat ID: {message.chat.id}, Chat Type: {message.chat.type}")
    if message.chat.type in ["group", "supergroup"]:
        bot.send_message(message.chat.id, f"Group Chat ID is `{message.chat.id}`", parse_mode="Markdown")
'''

