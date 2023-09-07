import discord
from discord.ext import commands
import os
import pandas as pd
from dotenv import load_dotenv
import sys

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
intents.members = True

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

def get_not_found_excel(data):
    df = pd.DataFrame(data)

    excel_file_name = "notFoundUsers.xlsx"
    excel_file = os.path.join(temp_dir, excel_file_name)
    df.to_excel(excel_file, index=False, engine="openpyxl")

    with open(excel_file, 'rb') as file:
        file_data = discord.File(file, filename=excel_file_name)
        return {
            "file": file,
            "discord_file": file_data,
            "excel_file": excel_file
            }

def print_progress(progress):
    sys.stdout.write("\r" + " " * 80)
    sys.stdout.write(f"\r{progress}")
    sys.stdout.flush()


# message event handler
@bot.command()
async def give_roles(ctx, role_name: str):
    if ctx.author == bot.user:
        return

    if ctx.message.attachments:
        for attachment in ctx.message.attachments:
            if attachment.filename.endswith((".xlsx", ".csv")):
                os.makedirs(temp_dir, exist_ok=True)
                file_path = os.path.join(temp_dir, attachment.filename)
                await attachment.save(file_path)

                data_objects = read_excel(file_path)

                if data_objects:
                    role_name = role_name.replace("-", " ")
                    role = discord.utils.get(ctx.guild.roles, name=role_name)
                    if role is None:
                        print("role does not exist")
                        return

                    print("giving roles...")
                    await ctx.channel.send(f"Giving roles started")
                    notFoundList = []
                    for index, data_object in enumerate(data_objects):
                        print_progress(f"{round(100 * ((index + 1) / len(data_objects)), 1)}% - {data_object['name']}")    
                        
                        identifier = data_object["discord id"]
                        user = discord.utils.get(ctx.guild.members, name=identifier)
                        
                        if user is None:
                            user = discord.utils.get(ctx.guild.members, id=identifier)

                        if user is None:
                            user = discord.utils.get(ctx.guild.members, mention=identifier)
                        
                        if user is None and "#" in identifier:
                            identifier = identifier.split("#")[0]
                            user = discord.utils.get(ctx.guild.members, name=identifier)

                            if user is None:
                                user = discord.utils.get(ctx.guild.members, id=identifier)

                            if user is None:
                                user = discord.utils.get(ctx.guild.members, mention=identifier)

                        if user is None:
                            notFoundList.append(data_object)
                            continue

                        await user.add_roles(role)

                    print("\nGiving roles finished")
                    if len(notFoundList) != 0:
                        notFoundUsersExcel = get_not_found_excel(notFoundList)
                        await ctx.channel.send(f"Giving roles finished\n{len(notFoundList)} users not found, they are provided in the attached file.", file=notFoundUsersExcel["discord_file"])
                        notFoundUsersExcel["file"].close()
                        os.remove(notFoundUsersExcel["excel_file"])

                    else:
                        await ctx.channel.send("Giving roles finished\nAll the users found!")

                os.remove(file_path)


# on start event handler
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user.name} ({bot.user.id})")
    print("----------------------------------------------")

    startup_channel_id = 776406395770765314
    startup_channel = bot.get_channel(startup_channel_id)

    if startup_channel:
        await startup_channel.send("Hey, I'm ready!")


# =================== main ===================
def main():
    bot.run(BOT_TOKEN)


if __name__ == "__main__":
    main()
