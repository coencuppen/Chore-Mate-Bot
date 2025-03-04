import os
import requests
import pandas as pd
from datetime import datetime
from telegram import Bot
import asyncio

def read_config(file_path):
    config = {}
    with open(file_path, 'r') as file:
        for line in file:
            # Split each line into key-value pair
            key, value = line.strip().split('=', 1)
            config[key] = value
    return config

# Load configuration
config = read_config("config.txt")

# Extract values from the config dictionary
BOT_TOKEN = config.get('BOT_TOKEN')
CHAT_ID = config.get('CHAT_ID')
huistakenURL = config.get('huistakenURL')
afwasRoosterURL = config.get('afwasRoosterURL')

# Step 2: Define the local file path
huistaakRoosterPath = "huistaakRooster.xlsx"
afwasRoosterPath = "afwasrooster.xlsx"




errorLog = []

# Step 3: Check if the file already exists
def getExcelFiles(url, path, forceDownload = False):
    if not os.path.exists(path) or forceDownload:
        # Step 4: Download the Excel file
        print(f"downlaoding url:{url}")
        response = requests.get(url)
        
        # Step 5: Save the file locally
        with open(path, "wb") as file:
            file.write(response.content)
        
        print(f'Downloaded: {path}')
    else:
        print(f'File already exists: {path}')

def capitalizeFirstLetter(s):
    if len(s) > 0 and not s[0].isupper():
        return s[0].upper() + s[1:]
    return s


def getTodaysDishWasher():
    # Read the Excel file
    df = pd.read_excel(afwasRoosterPath, engine='openpyxl')

    # Strip column names in case of leading/trailing spaces
    df.columns = df.columns.str.strip()

    # Ensure the 'DAG' column is treated as strings with stripped spaces
    df['DAG'] = df['DAG'].astype(str).str.strip()

    # Get today's date in the expected format
    today = datetime.now().strftime('%Y-%m-%d 00:00:00')

    # Filter the dataframe for today's date
    today_row = df[df['DAG'] == today]

    if not today_row.empty:
        # Get the name of the person
        person = today_row['PERSOON'].values[0]
        print(f"Today's dishwasher: {person}")
        return capitalizeFirstLetter(person)
    else:
        print("No entry found for today's date.")
    
    return None

def getTodaysTasks(file_path):
    # Load the Excel file into a DataFrame without headers
    df = pd.read_excel(file_path, engine='openpyxl', header=None, sheet_name=0)

    # Get today's date in the format that matches the DataFrame's column headers
    today = datetime.now().strftime('%Y-%m-%d 00:00:00')

    # Extract the date row (assumed to be the first row)
    date_row = df.iloc[0].astype(str)  # Convert all values to string for comparison

    # Find the column index for today's date
    today_index = date_row[date_row == today].index

    if today_index.empty:
        return None

    task_index = today_index[0]  # Get the first occurrence

    result = f"Taken voor vandaag:\n"

    # Iterate through all rows, starting from the second row (index 1) to get names and tasks
    for index, row in df.iloc[1:].iterrows():
        person = row.iloc[0]  # First column contains names
        task = row[task_index]  # Task for today's date
        result += f"{person}: {task}\n"

    result += "\nSucces met de taken iedereen! 🚀"

    return result.strip()

async def send_telegram_message(message):
    bot = Bot(token=BOT_TOKEN)
    await bot.send_message(chat_id=CHAT_ID, text=message)

def send_message_sync(message):
    asyncio.run(send_telegram_message(message))


def init():
    getExcelFiles(huistakenURL, huistaakRoosterPath, False)
    getExcelFiles(afwasRoosterURL, afwasRoosterPath, False)

if __name__ == '__main__':
    init()
    huistaken = getTodaysTasks(huistaakRoosterPath)
    if huistaken:
        send_message_sync(huistaken)

    send_message_sync(f"Onze afwasheld van vandaag is {getTodaysDishWasher()}. Zet hem op! 🍽️🧼")
    
