import calendar
import configparser
from datetime import datetime
import pandas as pd
import requests
from google.oauth2.service_account import Credentials
import gspread

# Authenticate and connect to Google Sheets
def connect_to_google_sheets(sheet_name, autocreds):
    creds = Credentials.from_service_account_file(
        autocreds,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    client = gspread.authorize(creds)
    try:
        return client.open(sheet_name)
    except gspread.SpreadsheetNotFound:
        print(f"Spreadsheet '{sheet_name}' not found.")
        return None


def get_events(num_events, schedule_df):
    """
    Parse the schedule DataFrame and return upcoming events.
    """
    today = pd.to_datetime(datetime.today().strftime('%Y-%m-%d'))

    schedule_df = schedule_df.copy()
    schedule_df.dropna(subset=['Date'], inplace=True)
    schedule_df['Date'] = pd.to_datetime(schedule_df['Date'], errors='coerce')
    schedule_df = schedule_df[schedule_df['Date'] > today]

    if not schedule_df.empty:
        schedule_df = schedule_df.iloc[:int(num_events)].reset_index(drop=True)

    return schedule_df


def send_teams(labmeeting_settings, msteam_settings, cal):
    location = labmeeting_settings['room']
    zoom_link = labmeeting_settings['zoom']
    holiday_vocab = set(labmeeting_settings['holiday_vocab'].split(", "))

    webhook_name = msteam_settings['webhookname']
    webhook_url = msteam_settings['webhookUrl']

    lines = []
    first = True

    for _, row in cal.iterrows():
        date = row['Date']
        topic = row['Type']
        member = row['Presenter(s)']
        
        formatted_date = date.strftime('%Y-%m-%d')

        if pd.isna(member) or topic in holiday_vocab:
            line = f"<strong>{formatted_date}</strong> <font color='red'>{topic} - {member}</font>"
        else:
            if first:
                line = (f"<strong>{formatted_date}</strong> {member} | {topic} (location <strong>{location}</strong> and <a href='{zoom_link}'>Meeting Link</a>)")
                first = False
            else:
                line = f"<strong>{formatted_date}</strong> {member} | {topic}"
        lines.append(line)

    # Simplified Payload for easy Power Automate parsing
    payload = {
        "title": "Upcoming Lab Meeting Schedule",
        "message_list": "\n\n".join(lines),
        "sender": webhook_name,
        "date_sent": datetime.today().strftime('%Y-%m-%d')
    }

    try:
        response = requests.post(webhook_url, json=payload, timeout=10)
        if response.status_code != 200:
            print(f"MS Workflow failed: {response.status_code} - {response.text}")
        else:
            print("Data sent successfully to Microsoft Workflow \n")

    except requests.exceptions.RequestException as e:
        print(f"Connection Error: {e}")



def main():
    config = configparser.ConfigParser(interpolation=None)
    config.read('cal_config.cfg')

    labmeeting_settings = config['labmeeting']
    msteam_settings = config['teams']
    num_events_teams = msteam_settings['maxevents']

    spreadsheet = connect_to_google_sheets(
        labmeeting_settings['googlesheet'],
        labmeeting_settings['autocreds']
    )

    if spreadsheet is None:
        return

    schedule_df = pd.DataFrame(
        spreadsheet.worksheet("Schedule").get_all_records()
    )

    if schedule_df.empty:
        print("No events found in the schedule.")
        return

    upcoming_events = get_events(num_events_teams, schedule_df)
    print(upcoming_events)

    send_teams(labmeeting_settings, msteam_settings, upcoming_events)


if __name__ == "__main__":
    main()
