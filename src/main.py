import discord
from discord.ext import commands
import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# temp files directory
temp_dir = os.path.join(".", "temp")

# bot token
BOT_TOKEN = os.getenv("DISCORD_BOT_TOKEN")

# intializing bot
intents = discord.Intents.default()
intents.typing = False
intents.presences = False
intents.message_content = True

bot = commands.Bot(command_prefix="!", intents=intents)


# ============== functions =================
def read_excel(file_path):
    try:
        if file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path)
        elif file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            raise ValueError(
                "Unsupported file format. Only Excel (.xlsx) and CSV (.csv) files are supported."
            )

        df.columns = df.columns.str.lower()
        required_columns = ["name", "team", "discord id"]
        if not all(column in df.columns for column in required_columns):
            raise ValueError(
                "The required columns 'name', 'team', and 'discord id' are missing in the file."
            )

        data_objects = df.to_dict(orient="records")

        return data_objects

    except Exception as e:
        print(str(e))
        return None


# message event handler
@bot.event
async def on_message(message):
    if message.author == bot.user:
        return

    if message.attachments:
        for attachment in message.attachments:
            if attachment.filename.endswith((".xlsx", ".csv")):
                os.makedirs(temp_dir, exist_ok=True)
                file_path = os.path.join(temp_dir, attachment.filename)
                await attachment.save(file_path)

                data_objects = read_excel(file_path)

                if data_objects:
                    for data_object in data_objects:
                        name = data_object["name"]
                        await message.channel.send(f"Name: {name}")

                os.remove(file_path)

    await bot.process_commands(message)


# on start event handler
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user.name} ({bot.user.id})")
    print("----------------------------------------------")

    startup_channel_id = 776406395770765314
    startup_channel = bot.get_channel(startup_channel_id)

    if startup_channel:
        await startup_channel.send("Bot is now online and ready to process files!")


# =================== main ===================
def main():
    bot.run(BOT_TOKEN)


if __name__ == "__main__":
    main()
