from datetime import datetime, timedelta
import calendar
import configparser
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import warnings

def is_holiday_thursday(date, holidays_df):
    """Check if given Thursday is a holiday."""
    
    # Check for hardcoded federal holidays (New Year's, Independence Day, Thanksgiving, and Christmas/New Year's)
    if (
        # New Year's (first Thursday in January)
        (date.month == 1 and date.weekday() == 3 and date.day <= 7) or
        # Independence Day (fixed to July 4th on Thursday)
        (date.month == 7 and date.day == 4 and date.weekday() == 3) or
        # Thanksgiving (fourth Thursday in November)
        (date.month == 11 and date.weekday() == 3 and date.day >= 22 and date.day <= 28) or
        # Christmas/New Year's (last Thursday in December)
        (date.month == 12 and date.weekday() == 3 and 
         date.day >= calendar.monthrange(date.year, date.month)[1] - 6)
    ):
        return True, "Federal Holiday/BCM observed"
    
    # Convert 'Date' column to datetime if not already
    holidays_df['Date'] = pd.to_datetime(holidays_df['Date'], errors='coerce')
    
    # Check for custom holidays from spreadsheet
    date_str = date.strftime('%Y-%m-%d')
    loaded_holidays = holidays_df['Date'].dt.strftime('%Y-%m-%d')

    # Ensure valid date format and check if the date is in the list of holidays
    if date_str in loaded_holidays.values:  
        date_match = holidays_df[loaded_holidays == date_str]
        holiday_name = date_match['Holiday'].iloc[0]  
        return True, holiday_name
        
    return False, ""

def next_thursday_on_or_after(date):
    """Return the next Thursday on or after the given date."""
    days_until_thursday = (3 - date.weekday()) % 7
    return date + timedelta(days=days_until_thursday)

def generate_schedule(spreadsheet, future_events_limit=16):
    """Generate the lab meeting schedule, alternating between Data and Journal Club presentations."""
    
    # Load data from sheets
    try:
        rotation_df = pd.DataFrame(spreadsheet.worksheet("Rotation").get_all_records())
        holidays_df = pd.DataFrame(spreadsheet.worksheet("Holidays").get_all_records())
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError("Required worksheet not found in the spreadsheet.")
    

    # Convert date columns
    rotation_df['Data date'] = pd.to_datetime(rotation_df['Data date'], errors='coerce')
    rotation_df['JC date'] = pd.to_datetime(rotation_df['JC date'], errors='coerce')

    # Ensure both date columns have at least one valid date
    if rotation_df['Data date'].isna().all() or rotation_df['JC date'].isna().all():
        raise ValueError("Both 'Data date' and 'JC date' must contain at least one valid date; earliest of the two dates starts cycle while the other serves as placeholder for start.")

    # Find the most recent date from both columns
    max_data_date = rotation_df['Data date'].max()
    max_jc_date = rotation_df['JC date'].max()

    # Get indices of the maximum dates
    data_index = rotation_df['Data date'].idxmax()
    jc_index = rotation_df['JC date'].idxmax()

    # Check if the difference between max dates is greater than 3 months
    if max_data_date is not pd.NaT and max_jc_date is not pd.NaT:
        date_diff = abs((max_data_date - max_jc_date).days)
        if date_diff > 90:
            warnings.warn(f"Warning: The maximum dates in 'Data date' and 'JC date' columns are more than 3 months apart ({date_diff} days).")

    start_date = min(max_data_date, max_jc_date)
    initial_jc = max_jc_date < max_data_date
    data_count = 0

    # get rotation lists
    rotation_data = rotation_df["Data rotation"].tolist()
    rotation_jc = rotation_df["JC rotation"].tolist()
    
    # Initialize schedule
    schedule = []
    current_date = start_date

    num_jc_presenters = 2 # get two presenters

    # Generate schedule
    while len(schedule) < future_events_limit:
        # Move to next Thursday
        current_date = next_thursday_on_or_after(current_date)
            
        # Check for holidays
        is_holiday, holiday_name = is_holiday_thursday(current_date, holidays_df)
        if is_holiday:
            schedule.append([
                current_date.strftime("%Y-%m-%d"),
                "Holiday",
                holiday_name
            ])
            current_date += timedelta(days=7)
            continue

        # Schedule next presentation
        if initial_jc and data_count == 0:
            # Initial Journal Club presentation
            presenters = rotation_jc[jc_index:jc_index + num_jc_presenters] 
            schedule.append([
                current_date.strftime("%Y-%m-%d"),
                "Journal Club",
                ", ".join(presenters)
            ])
            
            for presenter in presenters:
                rotation_df.loc[rotation_df["JC rotation"] == presenter, "JC date"] = current_date
            
            jc_index = (jc_index + num_jc_presenters) % len(rotation_jc)
            initial_jc = False
            
        elif data_count < 3:
            # Data presentation
            presenter = rotation_data[data_index % len(rotation_data)]
            schedule.append([
                current_date.strftime("%Y-%m-%d"),
                "Data",
                presenter
            ])
            
            rotation_df.loc[rotation_df["Data rotation"] == presenter, "Data date"] = current_date
            data_index = (data_index + 1) % len(rotation_data)
            data_count += 1
            
        else:
            # Journal Club presentation
            presenters = rotation_jc[jc_index:jc_index + num_jc_presenters]
            schedule.append([
                current_date.strftime("%Y-%m-%d"),
                "Journal Club",
                ", ".join(presenters)
            ])
            
            for presenter in presenters:
                rotation_df.loc[rotation_df["JC rotation"] == presenter, "JC date"] = current_date
            
            jc_index = (jc_index + num_jc_presenters) % len(rotation_jc)
            data_count = 0

        current_date += timedelta(days=7)

    # Update spreadsheets with the schedule
    try:
        schedule_sheet = spreadsheet.worksheet("Schedule")
        schedule_sheet.append_rows(schedule)
    except gspread.WorksheetNotFound:
        # Create the sheet if it doesn't exist
        schedule_sheet = spreadsheet.add_worksheet(title="Schedule", rows="100", cols="10")
        schedule_sheet.append_rows([["Date", "Type", "Presenter(s)"]] + schedule)

    # # Update the rotation sheet
    # rotation_df_update = rotation_df.copy()
    # for col in ['Data date', 'JC date']:
    #     rotation_df_update[col] = rotation_df_update[col].dt.strftime('%Y-%m-%d').fillna('')
    
    # try:
    #     rotation_sheet = spreadsheet.worksheet("Rotation")
    #     rotation_sheet.update([
    #         rotation_df_update.columns.values.tolist()
    #     ] + rotation_df_update.values.tolist())
    # except gspread.exceptions.WorksheetNotFound:
    #     print("Rotation sheet not found. Skipping update.")

    print(pd.DataFrame(spreadsheet.worksheet("Schedule").get_all_records()))
    return schedule, rotation_df

def main():
    # Load config
    config = configparser.ConfigParser()
    config.read('cal_config.cfg')
    labmeeting_settings = config['labmeeting']
    
    # Set up credentials and connect
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
        generate_schedule(spreadsheet, int(labmeeting_settings['schedule_events_count']))
    except gspread.SpreadsheetNotFound:
        print(f"Spreadsheet '{labmeeting_settings['googlesheet']}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
