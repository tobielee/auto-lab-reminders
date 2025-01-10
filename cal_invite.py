import calendar
from datetime import datetime
import configparser
import os

import argparse
import pandas as pd

import base64
from email.message import EmailMessage
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import gspread


def get_credentials():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/calendar",
        "https://www.googleapis.com/auth/gmail.send" 
    ]
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(labmeeting_settings['usercreds'], SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def init_services(creds):
    try:
        sheets = gspread.authorize(creds)
        calendar = build('calendar', 'v3', credentials=creds)
        gmail = build('gmail', 'v1', credentials=creds)  
        spreadsheet = sheets.open(labmeeting_settings['googlesheet'])
        return spreadsheet, calendar, gmail
    except gspread.SpreadsheetNotFound:
        print(f"Spreadsheet '{labmeeting_settings['googlesheet']}' not found.")
        return None, None, None
    except Exception as e:
        print(f"Error initializing services: {str(e)}")
        return None, None, None

def create_calendar_event(service, event_data, attendees):
    event_start = datetime.strptime(f"{event_data['date']} {labmeeting_settings['start_time']}", '%B %d, %Y %H:%M:%S')
    event_end = datetime.strptime(f"{event_data['date']} {labmeeting_settings['end_time']}", '%B %d, %Y %H:%M:%S')
    
    description = f"""
Hi Lab,

{event_data['presenter']} will be presenting {'data' if event_data['type'] == 'Data' else 'journal club articles'} at our next lab meeting.

Meeting will be held in {labmeeting_settings['room']} Breast center conference room and virtually at {labmeeting_settings['zoom']}.

If you have any questions regarding scheduling, let {labmeeting_settings['email']} know.

Best,
XZLab Bot

P.S. 

{zoom_extra_text}
"""

    event = {
        'summary': f"[XZ Lab Meeting]: {event_data['presenter']} | {event_data['type']}",
        'location': f"{labmeeting_settings['room']}, {labmeeting_settings['zoom']}",
        'description': description,
        'start': {'dateTime': event_start.isoformat(), 'timeZone': labmeeting_settings['timezone']},
        'end': {'dateTime': event_end.isoformat(), 'timeZone': labmeeting_settings['timezone']},
        'attendees': [{'email': email} for email in attendees],
        'reminders': {'useDefault': True},
        'visibility': 'default',
        'guestsCanSeeOtherGuests': True,
        'guestsCanInviteOthers': False,
        'guestsCanModify': False
    }

    try:
        service.events().insert(
            calendarId='primary',
            body=event,
            sendUpdates='all'
        ).execute()
        print(f"Calendar event created successfully")
        return True
    except HttpError as error:
        print(f"Calendar error: {error}")
        return False

def get_event(spreadsheet, exact_date=None):
    """
    Get lab meeting event. If exact_date is provided, look for event on that specific date.
    Otherwise, get the next upcoming event after today.
    """
    schedule = pd.DataFrame(spreadsheet.worksheet("Schedule").get_all_records())
    schedule['Date'] = pd.to_datetime(schedule['Date'])
    
    if exact_date:
        matches = schedule[
            (schedule['Date'].dt.date == exact_date.date())
        ]
    else:
        matches = schedule[
            (schedule['Date'].dt.date > datetime.now().date())
        ]
    
    matches = matches.sort_values('Date')
    
    if matches.empty:
        return None
        
    event = matches.iloc[0] # just get the first event for the (next) date with an event
    return {
        'date': event['Date'].strftime('%A %B %d, %Y'),
        'type': event['Type'],
        'presenter': event['Presenter(s)']
    }


def send_gmail(service, recipients, subject, body):
    """
    Sends an email to multiple recipients using the Gmail API.

    Args:
        service: Authorized Gmail API client.
        recipients: List of recipient email addresses.
        subject: Subject of the email.
        body: Body of the email.

    Returns:
        Boolean indicating success or failure.
    """
    try:
        # Create an email message
        message = EmailMessage()
        message.set_content(body)
        message["To"] = ", ".join(recipients)  # Combine recipients into a single string
        message["From"] = "your_email@example.com"  # Replace with your sender email
        message["Subject"] = subject

        # Encode the message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        # Create the payload for the Gmail API
        create_message = {"raw": encoded_message}

        # Send the email
        service.users().messages().send(
            userId="me", body=create_message
        ).execute()

        print("Email sent successfully")
        return True
    except HttpError as error:
        print(f"An error occurred: {error}")
        return False


def main():
    parser = argparse.ArgumentParser(description='Lab Meeting Calendar Manager')
    parser.add_argument('--auto', action='store_true',
                       help='Run in automated mode (looking exactly 7 days ahead)')
    args = parser.parse_args()

    config = configparser.ConfigParser()
    config.read('cal_config.cfg')
    global labmeeting_settings, zoom_extra_text
    labmeeting_settings = config['labmeeting']
    zoom_extra_text = labmeeting_settings['zoomextras']
    
    creds = get_credentials()
    spreadsheet, calendar, gmail = init_services(creds)
    if not all([spreadsheet, calendar, gmail]):
        print("Error initializing services.")
        return

    exact_date = datetime.now() + pd.Timedelta(days=7) if args.auto else None
    event_data = get_event(spreadsheet, exact_date)
    
    if not event_data:
        print("No upcoming events found.")
        return

    emails = [row['Email'] for row in spreadsheet.worksheet("Emails").get_all_records()]

    if event_data['type'] == 'Holiday' or event_data['type'] == 'Cancel':
        print(f"Sending email reminder for {event_data['type']} event on {event_data['date']}")
        send_gmail(
            service=gmail,
            recipients=emails,
            subject=f"[XZ Lab Meeting]: No lab meeting {event_data['date']}",
            body=f"""
Hi Lab,

Just a reminder: we will not have lab meeting on {event_data['date']} due to {event_data['presenter']}.

If you have any questions regarding scheduling, let {labmeeting_settings['email']} know.

Best,  
XZLab Bot
"""
            )
    else:
        create_calendar_event(calendar, event_data, emails)

if __name__ == "__main__":
    main()