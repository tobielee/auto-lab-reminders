from dataclasses import dataclass

import configparser
import argparse
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from typing import Optional

import os
import smtplib
import logging
from uuid import uuid4
from datetime import datetime, timedelta, date, timezone
from typing import List, Dict, Any
from zoneinfo import ZoneInfo  # Python 3.9+

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from icalendar import Calendar, Event, vCalAddress, vText

@dataclass
class MeetingConfig:
    """Holds all settings loaded from the config file."""
    start_time: str
    end_time: str
    timezone: str
    room: str
    zoom: str
    holiday_vocab: List[str]
    googlesheet: str
    autocreds_path: str
    auto_emailer: str
    contact_email: str
    smtp_server: str
    smtp_port: int
    # You can add other settings from your config here

@dataclass
class LabEvent:
    """Represents a single event from the schedule."""
    # Store the date as a proper date object. Formatting can happen later.
    event_date: date 
    event_type: str
    presenter: str

# Configure basic logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_config(filepath: str = 'cal_config.cfg') -> MeetingConfig:
    """Loads settings from a config file and returns a MeetingConfig object."""
    parser = configparser.ConfigParser()
    parser.read(filepath)

    if 'labmeeting' not in parser:
        raise ValueError("Config file must have a [labmeeting] section.")

    settings = parser['labmeeting']

    return MeetingConfig( # TODO: Add more fields as needed
        smtp_server=settings.get('smtp_server'),
        smtp_port=settings.getint('smtp_port', 587),
        googlesheet=settings.get('googlesheet'),
        start_time=settings.get('start_time'),
        end_time=settings.get('end_time'),
        timezone=settings.get('timezone'),
        room=settings.get('room'),
        zoom=settings.get('zoom'),
        holiday_vocab=settings.get('holiday_vocab', '').split(", "),
        autocreds_path=settings.get('autocreds'), 
        auto_emailer=settings.get('autoemailer'), 
        contact_email=settings.get('email')       
    )

def get_next_event(spreadsheet, exact_date: Optional[date] = None) -> Optional[LabEvent]:
    """
    Fetches the next event from the spreadsheet and returns an LabEvent object.
    
    Args:
        spreadsheet: The gspread spreadsheet object.
        exact_date: If provided, finds the event on this specific date.
                    Otherwise, finds the next upcoming event.

    Returns:
        An LabEvent object for the found event, or None if no event is found.
    """
    try:
        schedule_df = pd.DataFrame(spreadsheet.worksheet("Schedule").get_all_records())
        schedule_df['Date'] = pd.to_datetime(schedule_df['Date']).dt.date
    except Exception as e:
        logging.error(f"Error fetching or parsing spreadsheet data: {e}")
        return None

    if exact_date:
        # Filter for the exact date
        matches = schedule_df[schedule_df['Date'] == exact_date]
    else:
        # Filter for dates in the future
        today = datetime.now().date()
        matches = schedule_df[schedule_df['Date'] > today]

    if matches.empty:
        return None

    # Sort by date to ensure we get the soonest one
    event_row = matches.sort_values('Date').iloc[0]

    # Map DataFrame columns to our LabEvent object
    return LabEvent(
        event_date=event_row['Date'],
        event_type=event_row['Type'],
        presenter=event_row["Presenter(s)"]
    )


def send_gmail_smtp(
    sender_email: str,
    recipients: List[str],
    subject: str,
    body: str,
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587
) -> bool:
    password = os.getenv("SENDER_APP_PASSWORD")
    if not password:
        logging.error("Missing SENDER_APP_PASSWORD environment variable.")
        return False

    try:
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = sender_email
        msg["To"] = ", ".join(recipients)

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, recipients, msg.as_string())

        logging.info("Email sent successfully via SMTP.")
        return True
    except Exception as e:
        logging.error(f"SMTP email error: {e}")
        return False


def send_calendar_invite_smtp(
    sender_email: str,
    recipients: List[str],
    subject: str,
    event_data: Dict[str, Any],
    settings: Dict[str, Any],
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587
) -> bool:
    """
    Sends an RFC-compliant calendar invite via SMTP using an App Password.
    """
    password = os.getenv("SENDER_APP_PASSWORD")
    if not password:
        logging.error("Missing SENDER_APP_PASSWORD environment variable.")
        return False

    try:
        # Parse date and time
        tz = ZoneInfo(settings['timezone'])
        event_start_naive = datetime.combine(event_data['date'], datetime.strptime(settings['start_time'], "%H:%M:%S").time())
        event_end_naive = datetime.combine(event_data['date'], datetime.strptime(settings['end_time'], "%H:%M:%S").time())
        event_start = event_start_naive.replace(tzinfo=tz)
        event_end = event_end_naive.replace(tzinfo=tz)
    except (ValueError, KeyError) as e:
        logging.error(f"Date/time parsing error: {e}")
        return False

    # Build the iCalendar event
    cal = Calendar()
    cal.add('prodid', '-//XZLab//Lab Meeting Scheduler//EN')
    cal.add('version', '2.0')
    cal.add('method', 'REQUEST')

    event = Event()
    event.add('summary', subject)
    event.add('dtstart', event_start)
    event.add('dtend', event_end)
    event.add('dtstamp', datetime.now(timezone.utc))
    event.add('uid', str(uuid4()))
    event.add('location', vText(f"{settings['room']}, {settings['zoom']}"))
    event.add('description', vText(event_data['description']))

    event['dtstart'].params['tzid'] = vText(settings['timezone'])
    event['dtend'].params['tzid'] = vText(settings['timezone'])

    # Add organizer
    organizer = vCalAddress(f'mailto:{sender_email}')
    organizer.params['cn'] = vText("XZLab Bot")
    event.add('organizer', organizer)

    # Add attendees
    for email in recipients:
        attendee = vCalAddress(f'mailto:{email}')
        attendee.params['cn'] = vText(email)
        attendee.params['ROLE'] = vText('REQ-PARTICIPANT')
        attendee.params['RSVP'] = vText('TRUE')
        event.add('attendee', attendee, encode=0)

    cal.add_component(event)

    # # Build email with calendar attachment
    # msg = MIMEMultipart('alternative')
    # msg['Subject'] = subject
    # msg['From'] = f"XZLab Bot <{sender_email}>"
    # msg['To'] = ", ".join(recipients)

    # msg.attach(MIMEText(event_data['description'], 'plain'))

    # cal_part = MIMEText(cal.to_ical().decode(), 'calendar', _charset='utf-8')
    # cal_part.set_param('method', 'REQUEST')
    # cal_part.add_header('Content-Disposition', 'attachment', filename='invite.ics')
    # msg.attach(cal_part)

    # Build email with calendar attachment
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = f"XZLab Bot <{sender_email}>"
    msg['To'] = ", ".join(recipients)

    # Add this header so Outlook recognizes the email as a calendar invite
    msg.add_header('Content-Class', 'urn:content-classes:calendarmessage')

    # Plain text body
    msg.attach(MIMEText(event_data['description'], 'plain'))

    # Calendar part — note method param and content-disposition
    cal_part = MIMEText(cal.to_ical().decode(), 'calendar;method=REQUEST', _charset='utf-8')
    cal_part.add_header('Content-Disposition', 'inline; filename="invite.ics"')
    msg.attach(cal_part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, recipients, msg.as_string())
        logging.info("Calendar invite sent to %d attendees.", len(recipients))
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP auth failed. Check sender email and SENDER_APP_PASSWORD.")
        return False
    except Exception as e:
        logging.error(f"SMTP error sending calendar invite: {e}")
        return False


def get_service_account_credentials(path):
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    return Credentials.from_service_account_file(path, scopes=SCOPES)



def init_services(creds, spreadsheet_name: str):
    """
    Authorizes with the Google Sheets API and opens a specific spreadsheet.

    Args:
        creds: The authorized service account credentials.
        spreadsheet_name: The name of the Google Sheet to open.

    Returns:
        A gspread.Spreadsheet object if successful, otherwise None.
    """
    if not spreadsheet_name:
        logging.error("No spreadsheet name provided. Check your config file.")
        return None
    try:
        sheets_client = gspread.authorize(creds)
        spreadsheet = sheets_client.open(spreadsheet_name)
        logging.info(f"Successfully connected to spreadsheet: '{spreadsheet_name}'")
        return spreadsheet
    except gspread.SpreadsheetNotFound:
        logging.error(f"Spreadsheet '{spreadsheet_name}' not found. "
                      "Check the name and ensure the service account has access.")
        return None
    except Exception as e:
        logging.error(f"An unexpected error occurred while initializing gspread services: {e}")
        return None


def handle_holiday_event(event: LabEvent, attendees: list, config: MeetingConfig):
    """Sends a plain text email for a holiday or non-meeting event."""
    logging.info(f"LabEvent is a holiday/break: '{event.presenter}'. Sending a reminder email.")
    
    event_date_str = event.event_date.strftime('%A, %B %d')
    subject = f"[XZ Lab Meeting]: No lab meeting on {event_date_str}"
    body = f"""Hi Lab,

Just a reminder: we will not have lab meeting on {event_date_str} due to {event.presenter}.

If you have any questions regarding scheduling, let {config.contact_email} know.

Best,
XZLab Bot
"""
    # The sending function should get the password from env variables itself
    success = send_gmail_smtp(
        sender_email=config.auto_emailer,
        recipients=attendees,
        subject=subject,
        body=body
    )
    if success:
        logging.info("Holiday reminder email sent successfully.")
    else:
        logging.error("Failed to send holiday reminder email.")


def handle_regular_meeting(event: LabEvent, attendees: list, config: MeetingConfig):
    """Prepares and sends a calendar invite for a regular meeting."""
    logging.info(f"LabEvent is a regular meeting with '{event.presenter}'. Sending a calendar invite.")
    # Prepare the data structures for the sending function
    event_data = {
        "presenter": event.presenter,
        "type": event.event_type,
        "date": event.event_date,
        "description": f"""
Hi Lab,

{event.presenter} will be presenting {'data' if event.event_type== 'Data' else 'journal club articles'} at our next lab meeting.

Meeting will be held in {config.room} Breast center conference room and virtually at {config.zoom}.

If you have any questions regarding scheduling, let {config.contact_email} know.

Best,
XZLab Bot
"""
    }

    settings_dict = {
        'start_time': config.start_time,
        'end_time': config.end_time,
        'timezone': config.timezone,
        'room': config.room,
        'zoom': config.zoom
    }

    success = send_calendar_invite_smtp(
        sender_email=config.auto_emailer,
        event_data=event_data,
        subject = f"[XZ Lab Meeting]: {event_data['presenter']} | {event_data['type']}",
        settings=settings_dict,
        recipients=attendees
        
    )
    if success:
        logging.info("Calendar invite sent successfully.")
    else:
        logging.error("Failed to send calendar invite.")

def main():
    parser = argparse.ArgumentParser(description='Lab Meeting Calendar Manager')
    parser.add_argument('--auto', action='store_true',
                        help='Run in automated mode (looking exactly 7 days ahead)')
    args = parser.parse_args()

    try:
        # 1. Load configuration into a clean object
        meeting_config = load_config('cal_config.cfg')

        # 2. Authenticate and initialize services using the config object
        creds = get_service_account_credentials(meeting_config.autocreds_path)
        spreadsheet = init_services(creds, meeting_config.googlesheet)
        if not spreadsheet:
            logging.error("Error initializing spreadsheet.")
            return

        # 3. Determine which date to look for based on --auto flag
        target_date = None
        if args.auto:
            target_date = datetime.now().date() + timedelta(days=7)
            logging.info(f"Running in auto mode. Looking for event on {target_date.strftime('%Y-%m-%d')}")

        # 4. Get the event for the target date (or the next upcoming one if not auto)
        event = get_next_event(spreadsheet, exact_date=target_date)

        if not event:
            if target_date:
                logging.info(f"No event found for {target_date.strftime('%Y-%m-%d')}.")
            else:
                logging.info("No upcoming events found in the schedule.")
            return

        logging.info(f"Found event: '{event.presenter}' on {event.event_date.strftime('%Y-%m-%d')}")
        # 5. Get the list of attendees
        attendees = [row['Email'] for row in spreadsheet.worksheet("Emails").get_all_records()]
        if not attendees:
            logging.warning("No attendees found in the 'Emails' sheet.")
        # 6. Decide whether to send a holiday reminder or a calendar invite
        if event.event_type in meeting_config.holiday_vocab:
            handle_holiday_event(event, attendees, meeting_config)
        else:
            handle_regular_meeting(event, attendees, meeting_config)
    except FileNotFoundError:
        logging.error("Error: The config file 'cal_config.cfg' was not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred in the main process: {e}")


if __name__ == "__main__":
    main()
