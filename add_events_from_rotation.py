from datetime import datetime, timedelta
import calendar
import configparser
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# --- Helper Functions ---

def get_next_thursday(date):
    """Return the next Thursday after the given date."""
    days_ahead = 3 - date.weekday() # Thursday is 3
    if days_ahead <= 0: # Target day already happened this week
        days_ahead += 7
    return date + timedelta(days=days_ahead)

def is_holiday_thursday(date, holidays_df):
    """Check if given Thursday is a holiday."""
    # 1. Check Hardcoded Federal Holidays
    if (
        (date.month == 1 and date.weekday() == 3 and date.day <= 7) or # New Year's
        (date.month == 7 and date.day == 4 and date.weekday() == 3) or # July 4th
        (date.month == 11 and date.weekday() == 3 and 22 <= date.day <= 28) or # Thanksgiving
        (date.month == 12 and date.weekday() == 3 and date.day >= 25) # Xmas/New Years week
    ):
        return True, "Federal Holiday/BCM observed"
    
    # 2. Check Spreadsheet Holidays
    if not holidays_df.empty:
        # Ensure dates are strings for comparison
        holidays_df['Date'] = pd.to_datetime(holidays_df['Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        date_str = date.strftime('%Y-%m-%d')
        
        if date_str in holidays_df['Date'].values:
            match = holidays_df[holidays_df['Date'] == date_str]
            return True, match['Holiday'].iloc[0]
        
    return False, ""

def get_next_presenter_index(history_df, event_type, rotation_list):
    """
    Looks at the schedule history to find the last person who presented
    and returns the index of the NEXT person in the rotation list.
    """
    if history_df.empty:
        return 0

    # Filter history for this specific event type (Data or JC)
    type_history = history_df[history_df['Type'] == event_type]
    
    if type_history.empty:
        return 0

    # Get the last presenter(s) string
    last_presenter_str = type_history.iloc[-1]['Presenter(s)']
    
    # Logic for Journal Club (often comma separated pairs)
    # We just need to find the last person in the pair in our list
    last_presenters = [x.strip() for x in last_presenter_str.split(',')]
    last_person = last_presenters[-1] # Take the last name listed

    try:
        current_index = rotation_list.index(last_person)
        return (current_index + 1) % len(rotation_list)
    except ValueError:
        print(f"Warning: Last presenter '{last_person}' not found in rotation list. Starting from top.")
        return 0

def get_cycle_state(history_df):
    """
    Determines how many 'Data' meetings have happened since the last 'Journal Club'.
    Returns an integer (0, 1, or 2).
    """
    if history_df.empty:
        return 0
    
    # We reverse the dataframe to look backwards
    reversed_df = history_df.iloc[::-1]
    
    data_count = 0
    for _, row in reversed_df.iterrows():
        evt_type = row['Type']
        if evt_type == "Journal Club":
            break # Stop counting when we hit a JC
        elif evt_type == "Data":
            data_count += 1
        # Ignore Holidays
            
    return data_count

# --- Main Logic ---

def generate_schedule(spreadsheet, future_events_limit=16, dry_run=False):
    
    # 1. Load Data
    try:
        rotation_sheet = spreadsheet.worksheet("Rotation")
        rotation_df = pd.DataFrame(rotation_sheet.get_all_records())
        
        # Get clean lists of names, ignoring empty strings
        rotation_data = [x for x in rotation_df["Data rotation"].tolist() if x]
        rotation_jc = [x for x in rotation_df["JC rotation"].tolist() if x]
        
        holidays_sheet = spreadsheet.worksheet("Holidays")
        holidays_df = pd.DataFrame(holidays_sheet.get_all_records())
        
        # Open Schedule or create if missing
        try:
            schedule_sheet = spreadsheet.worksheet("Schedule")
            schedule_data = schedule_sheet.get_all_records()
            schedule_df = pd.DataFrame(schedule_data)
        except gspread.WorksheetNotFound:
            schedule_sheet = spreadsheet.add_worksheet(title="Schedule", rows="100", cols="10")
            schedule_sheet.append_row(["Date", "Type", "Presenter(s)"])
            schedule_df = pd.DataFrame(columns=["Date", "Type", "Presenter(s)"])

    except gspread.exceptions.WorksheetNotFound as e:
        raise ValueError(f"Required worksheet not found: {e}")

    # 2. Analyze History to determine start state
    
    # Determine Start Date
    if not schedule_df.empty and 'Date' in schedule_df.columns:
        # Convert to datetime to find max
        dates = pd.to_datetime(schedule_df['Date'], errors='coerce')
        last_date = dates.max()
        if pd.isna(last_date):
            start_date = datetime.now()
        else:
            start_date = last_date
    else:
        start_date = datetime.now()

    # Determine Indices
    data_index = get_next_presenter_index(schedule_df, "Data", rotation_data)
    # Note: For JC we might jump by 2s, but we find the index of the *next* person
    jc_index = get_next_presenter_index(schedule_df, "Journal Club", rotation_jc)
    
    # Determine Cycle (0, 1, or 2 Data meetings have passed)
    data_streak = get_cycle_state(schedule_df)

    # 3. Generate New Events
    new_events = []
    current_date = start_date
    
    # Params
    NUM_JC_PRESENTERS = 2
    
    while len(new_events) < future_events_limit:
        current_date = get_next_thursday(current_date)
        
        # Check Holiday
        is_holiday, holiday_name = is_holiday_thursday(current_date, holidays_df)
        if is_holiday:
            new_events.append([
                current_date.strftime("%Y-%m-%d"),
                "Holiday",
                holiday_name
            ])
            continue # Skip to next loop iteration (date stays same, loop finds next thursday)

        # Logic: 3 Data -> 1 JC
        # If data_streak is 0, 1, or 2 -> Schedule Data. 
        # If data_streak >= 3 -> Schedule JC, reset streak.
        
        if data_streak < 3:
            # --- DATA PRESENTATION ---
            presenter = rotation_data[data_index]
            
            new_events.append([
                current_date.strftime("%Y-%m-%d"),
                "Data",
                presenter
            ])
            
            # Update state
            data_index = (data_index + 1) % len(rotation_data)
            data_streak += 1
            
        else:
            # --- JOURNAL CLUB ---
            # Grab next N presenters
            presenters = []
            temp_idx = jc_index
            for _ in range(NUM_JC_PRESENTERS):
                presenters.append(rotation_jc[temp_idx])
                temp_idx = (temp_idx + 1) % len(rotation_jc)
            
            new_events.append([
                current_date.strftime("%Y-%m-%d"),
                "Journal Club",
                ", ".join(presenters)
            ])
            
            # Update state
            jc_index = temp_idx # Update global index to where we left off
            data_streak = 0 # Reset cycle

    # 4. Append to Google Sheet
    if new_events:
        if dry_run:
            print("\n--- DRY RUN PREVIEW ---")
            print(f"{'Date':<12} | {'Type':<15} | {'Presenter(s)'}")
            print("-" * 45)
            for event in new_events:
                print(f"{event[0]:<12} | {event[1]:<15} | {event[2]}")
            print("--- END PREVIEW (No changes made to Google Sheets) ---\n")
        else:
            schedule_sheet.append_rows(new_events)
            print(f"Success! Appended {len(new_events)} new events starting from {new_events[0][0]}.")
    else:
        print("No events generated.")

def main():
    # Load config
    config = configparser.ConfigParser()
    config.read('cal_config.cfg')
    
    try:
        labmeeting_settings = config['labmeeting']
    except KeyError:
        print("Error: 'labmeeting' section not found in cal_config.cfg")
        return

    # Set up credentials
    creds = Credentials.from_service_account_file(
        labmeeting_settings['autocreds'],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    client = gspread.authorize(creds)
    
    try:
        spreadsheet = client.open(labmeeting_settings['googlesheet'])
        generate_schedule(spreadsheet, int(labmeeting_settings.get('schedule_events_count', 16)), dry_run=True)
    except gspread.SpreadsheetNotFound:
        print(f"Spreadsheet '{labmeeting_settings['googlesheet']}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()