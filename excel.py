import os
import requests
import pandas as pd
from datetime import datetime, timedelta
from telegram import Bot
import asyncio
import quote

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
    # Lees het Excel-bestand
    df = pd.read_excel(afwasRoosterPath, engine='openpyxl')

    # Strip kolomnamen voor het geval er spaties zijn
    df.columns = df.columns.str.strip()

    # Controleer of de kolom 'DAG' een datetime-object is, zo niet, converteer het
    if not pd.api.types.is_datetime64_any_dtype(df['DAG']):
        df['DAG'] = pd.to_datetime(df['DAG'], errors='coerce', dayfirst=True)

    # Haal de datum van vandaag op en zet het in hetzelfde formaat als de DataFrame
    today = pd.to_datetime(datetime.today().date())

    # Filter de DataFrame op de datum van vandaag
    today_row = df[df['DAG'] == today]

    if not today_row.empty:
        # Haal de naam van de persoon op
        person = today_row.iloc[0]['PERSOON']
        print(f"Today's dishwasher: {person}")
        return person.strip().capitalize()  # Zorgt voor nette hoofdletters
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

    result += "\nSucces met de taken iedereen! 🚀\n"

    result += quote.getRandomQuote()

    return result.strip()

async def send_telegram_message(message):
    bot = Bot(token=BOT_TOKEN)
    await bot.send_message(chat_id=CHAT_ID, text=message, parse_mode="HTML")

def send_message_sync(message):
    asyncio.run(send_telegram_message(message))

def check_and_extend_schedule():
    df = pd.read_excel(afwasRoosterPath, engine='openpyxl')
    df['DAG'] = pd.to_datetime(df['DAG'], errors='coerce')
    last_date = df['DAG'].max()
    seven_days_from_now = datetime.now() + timedelta(days=21)
    
    if pd.isna(last_date) or last_date < seven_days_from_now:
        new_dates = [(last_date + timedelta(days=i)).strftime('%Y-%m-%d 00:00:00') for i in range(1, 31)]
        
        
        
        print("Extended schedule with 30 new dates.")


def init():
    getExcelFiles(huistakenURL, huistaakRoosterPath, True)
    getExcelFiles(afwasRoosterURL, afwasRoosterPath, True)
    check_and_extend_schedule()

if __name__ == '__main__':
    init()
    huistaken = getTodaysTasks(huistaakRoosterPath)
    if huistaken:
        send_message_sync(huistaken)


    print("Afwasheld van vandaag:")
    print(getTodaysDishWasher())
    send_message_sync(f"<b>{getTodaysDishWasher()}</b> is onze afwasheld van vandaag. Zet hem op! 🍽️🧼")
    
