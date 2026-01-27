from dataclasses import dataclass
import configparser
import argparse
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from typing import Optional

import os
from pathlib import Path
from dotenv import load_dotenv

try:
    REPO_ROOT = Path(__file__).parent
except NameError:
    REPO_ROOT = Path.cwd()

ENV_PATH = REPO_ROOT / ".env"
if load_dotenv(dotenv_path=ENV_PATH):
    print(f"Loaded .env from {ENV_PATH}")
else:
    print(f"Failed to load .env from {ENV_PATH}")

import smtplib
import logging
from uuid import uuid4
from datetime import datetime, timedelta, date, timezone
from typing import List, Dict, Any
from zoneinfo import ZoneInfo  # Python 3.9+

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from icalendar import Calendar, Event, vCalAddress, vText
from email.utils import formatdate, make_msgid
import time

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
    contact_email: str
    smtp_server: str
    smtp_port: int
    batch_size: int # Added batch size field

@dataclass
class LabEvent:
    """Represents a single event from the schedule."""
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

    return MeetingConfig(
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
        contact_email=settings.get('email'),
        batch_size=settings.getint('batch_size', 1) # Added batch_size loader
    )

def chunk_recipients(recipient_list: list, batch_size: int):
    """Yield successive n-sized chunks from recipient_list."""
    for i in range(0, len(recipient_list), batch_size):
        yield recipient_list[i:i + batch_size]

def get_next_event(spreadsheet, exact_date: Optional[date] = None) -> Optional[LabEvent]:
    try:
        schedule_df = pd.DataFrame(spreadsheet.worksheet("Schedule").get_all_records())
        schedule_df['Date'] = pd.to_datetime(schedule_df['Date']).dt.date
    except Exception as e:
        logging.error(f"Error fetching or parsing spreadsheet data: {e}")
        return None

    if exact_date:
        matches = schedule_df[schedule_df['Date'] == exact_date]
    else:
        today = datetime.now().date()
        matches = schedule_df[schedule_df['Date'] > today]

    if matches.empty:
        return None

    event_row = matches.sort_values('Date').iloc[0]
    return LabEvent(
        event_date=event_row['Date'],
        event_type=event_row['Type'],
        presenter=event_row["Presenter(s)"]
    )


def send_gmail_smtp(
    recipients: List[str],
    subject: str,
    body: str,
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587
) -> bool:
    sender_email = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASSWORD")
    if not sender_email or not password:
        logging.error("Missing email credentials in environment.")
        return False

    try:
        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = f"XZLab Bot <{sender_email}>"
        # BCC Logic: Address the email 'To' the bot itself
        msg["To"] = sender_email 
        msg["Message-ID"] = make_msgid()
        msg["Date"] = formatdate(localtime=True)
        msg["User-Agent"] = "Mozilla Thunderbird"
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            # server.sendmail actually delivers to the recipients list
            server.sendmail(sender_email, recipients, msg.as_string())

        return True
    except Exception as e:
        logging.error(f"SMTP email error: {e}")
        return False


def send_calendar_invite_smtp(
    recipients: List[str],
    subject: str,
    event_data: Dict[str, Any],
    settings: Dict[str, Any],
    smtp_server: str = "smtp.gmail.com",
    smtp_port: int = 587
) -> bool:
    sender_email = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASSWORD")
    if not sender_email or not password:
        logging.error("Missing email credentials in environment.")
        return False

    try:
        tz = ZoneInfo(settings['timezone'])
        event_start = datetime.combine(event_data['date'], datetime.strptime(settings['start_time'], "%H:%M:%S").time()).replace(tzinfo=tz)
        event_end = datetime.combine(event_data['date'], datetime.strptime(settings['end_time'], "%H:%M:%S").time()).replace(tzinfo=tz)
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

    organizer = vCalAddress(f'mailto:{sender_email}')
    organizer.params['cn'] = vText("XZLab Bot")
    event.add('organizer', organizer)

    # Add attendees to ICS metadata for this batch
    for email in recipients:
        attendee = vCalAddress(f'mailto:{email}')
        attendee.params['cn'] = vText(email)
        attendee.params['ROLE'] = vText('REQ-PARTICIPANT')
        attendee.params['RSVP'] = vText('TRUE')
        event.add('attendee', attendee, encode=0)

    cal.add_component(event)

    # Build email
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = f"XZLab Bot <{sender_email}>"
    msg['To'] = sender_email # BCC Logic
    msg['Message-ID'] = make_msgid()
    msg['Date'] = formatdate(localtime=True)
    msg['User-Agent'] = "Microsoft Outlook 16.0"
    msg.add_header('Content-Class', 'urn:content-classes:calendarmessage')

    msg.attach(MIMEText(event_data['description'], 'plain'))
    ical_data = cal.to_ical().decode('utf-8')
    cal_part = MIMEText(ical_data, _subtype="calendar", _charset="utf-8")
    cal_part.replace_header("Content-Type", 'text/calendar; charset="utf-8"; method=REQUEST')
    cal_part.add_header("Content-Disposition", 'inline; filename="invite.ics"')
    msg.attach(cal_part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, recipients, msg.as_string())
        return True
    except Exception as e:
        logging.error(f"SMTP error sending calendar invite: {e}")
        return False


def get_service_account_credentials(path):
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    return Credentials.from_service_account_file(path, scopes=SCOPES)


def init_services(creds, spreadsheet_name: str):
    if not spreadsheet_name:
        logging.error("No spreadsheet name provided.")
        return None
    try:
        sheets_client = gspread.authorize(creds)
        spreadsheet = sheets_client.open(spreadsheet_name)
        logging.info(f"Successfully connected to spreadsheet: '{spreadsheet_name}'")
        return spreadsheet
    except Exception as e:
        logging.error(f"Error initializing gspread: {e}")
        return None


def handle_holiday_event(event: LabEvent, attendees: list, config: MeetingConfig):
    """Sends batched holiday reminder emails using BCC."""
    event_date_str = event.event_date.strftime('%A, %B %d')
    subject = f"[XZ Lab Meeting]: No lab meeting on {event_date_str}"
    body = f"""Hi Lab,

Just a reminder: we will not have lab meeting on {event_date_str} due to {event.presenter}.

If you have any questions regarding scheduling, let {config.contact_email} know.

Best,
XZLab Bot
"""
    batches = list(chunk_recipients(attendees, config.batch_size))
    for i, batch in enumerate(batches):
        logging.info(f"Sending holiday batch {i+1}/{len(batches)} ({len(batch)} recipients)")
        send_gmail_smtp(
            recipients=batch,
            subject=subject,
            body=body,
            smtp_server=config.smtp_server,
            smtp_port=config.smtp_port
        )
        time.sleep(2) # Delay for delivery reliability


def handle_regular_meeting(event: LabEvent, attendees: list, config: MeetingConfig):
    """Sends batched calendar invites using BCC."""
    event_data = {
        "presenter": event.presenter,
        "type": event.event_type,
        "date": event.event_date,
        "description": f"""
Hi Lab,

{event.presenter} will be presenting {'data' if event.event_type== 'Data' else 'journal club articles'} at our next lab meeting.

Meeting will be held in {config.room} and virtually at {config.zoom}.

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

    batches = list(chunk_recipients(attendees, config.batch_size))
    for i, batch in enumerate(batches):
        logging.info(f"Sending invite batch {i+1}/{len(batches)} ({len(batch)} recipients)")
        success = send_calendar_invite_smtp(
            event_data=event_data,
            subject = f"[XZ Lab Meeting]: {event_data['presenter']} | {event_data['type']}",
            settings=settings_dict,
            recipients=batch,
            smtp_server=config.smtp_server,
            smtp_port=config.smtp_port
        )
        if success:
            logging.info(f"Batch {i+1} sent successfully.")
        time.sleep(2)


def main():
    parser = argparse.ArgumentParser(description='Lab Meeting Calendar Manager')
    parser.add_argument('--auto', action='store_true', help='Search 7 days ahead')
    args = parser.parse_args()

    try:
        meeting_config = load_config('cal_config.cfg')
        creds = get_service_account_credentials(meeting_config.autocreds_path)
        spreadsheet = init_services(creds, meeting_config.googlesheet)
        if not spreadsheet: return

        target_date = (datetime.now().date() + timedelta(days=7)) if args.auto else None
        event = get_next_event(spreadsheet, exact_date=target_date)

        if not event:
            logging.info("No upcoming events found.")
            return

        attendees = [row['Email'] for row in spreadsheet.worksheet("Emails").get_all_records() if row.get('Email')]
        
        if event.event_type in meeting_config.holiday_vocab:
            handle_holiday_event(event, attendees, meeting_config)
        else:
            handle_regular_meeting(event, attendees, meeting_config)
    except Exception as e:
        logging.error(f"Unexpected error: {e}")

if __name__ == "__main__":
    main()