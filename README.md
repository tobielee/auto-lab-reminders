# auto-lab-reminders

This Python application allows users to set schedules in Google Sheets and send Google Calendar invites (emails) and MS Teams reminders based on the schedule. Accessing calendar data is done through Google API. Please read https://developers.google.com/identity/protocols/oauth2 to understand how access tokens work to use Google API. 

## Requirements
- A gmail account (to use gmail, google sheets, google calendar)
  
- **Python Version**: Python >= 3.8  
  - *Recommendation*: If you don't already have a valid version of Python, consider installing the latest version.  

## Python Dependencies  
To install the required Python libraries, use the following command:  
```bash
pip install argparse pandas google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread pymsteams icalendar
```

## Components  

The application consists of the following:  

### Google Sheets 

The main Google Sheets contains the following sheets:  

1. **Rotation**  
   Specifies the rotation logic for scheduling events or tasks.  

2. **Emails**  
   Handles sending email invites to participants based on the schedule.  

3. **Holidays**  
   - Default holidays are hardcoded based on BCM holidays (falling on Thursdays) and include a 2-week winter break (which don't need to appear on this sheet): 
     1. first Thursday in Jan
     2. if July 4 falls on Thursday
     3. fourth Thursday in November
     4. last Thursday in December  
   - This sheet is for labeling  unaccounted for "holidays" or more specifically mark Shawn's absences that are not BCM holidays (may consider relabeling this sheet in the future, since holiday might not be most apt) 

4. **Schedule**
   
   - The most important tracking spreadsheet with three columns: Date, Type, Presenter(s). Each row represents an individual event.
   
   &nbsp;

   [See Future Features](#futurefeats)
   - **Autogenerate Schedule**: Automatically creates a schedule based on the inputs.  
   - **Autoupdate Schedule**: Updates the schedule dynamically to reflect changes.  

### Configuration File  

- The config file is used to define settings for calendar invites and Teams reminders.  
- Allows customization of reminder timings and other notification preferences.

For sending Gmail, authorization token is required, so you must generate this (currently program expects token as .json) to be able to send emails. 
Authorizing Oauth2 account is set in the `usercreds` .json which you can create and download from https://console.cloud.google.com/apis/credentials

For all other purposes (editing Google sheet/calendar), a service account set in `autocreds` is used (which doesn't require a token to be generated).

I think I could have just used usercreds for both cases rather than having the complexity of the two but maybe to hedge against having to recreate a token occasionally I have autocreds set..

For MS Teams notifications a webhook is used and set in the `webhookname` and `webhookUrl`:  
https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook?tabs=newteams%2Cdotnet

Other fields in config file include:

[labmeeting]
- `googlesheet`: name of google sheet on google drive
- `room`: meeting location room
- `zoom`: Zoom link
- `email`: contact email for managing calendar (app)
- `start_time`: start of meeting
- `end_time`: end of meeting
- `timezone`: timezone for meeting
- `schedule_envents_count`: number of events (rows) to add to calendar with `generate_schedule.py`
- `holiday_vocab`: comma separated list of vocab for indicating what is a holiday based on what's given in the Type column of the `Schedule` tab of the Google sheets 
- `zoomextras`: additional text appended onto calendar invite message/description 

[teams]
- `maxevents`: number of events to show on MS teams notification

### Scripts  

The application includes two Python scripts for execution, which can be run manually or scheduled via a cron job:  

1. **`cal_invite.py`**  
   - Handles sending calendar invites based on the schedule. This will send a email for instances where lab meeting has been canceled due to Shawn's absence or other circumstances.
   - Manual runs of the script get the next proximal event, while including `--auto` flag in a cron job triggers invites only for events a week away (hardcoded).

Helper functions `get_token.py` and `refresh_token.py` should be used to get and retain active token to send email/calendar invite. 

2. **`msteams_remind.py`**  
   - Sends notifications via Microsoft Teams based on the schedule.  

3. (optional) **`generate_schedule.py`** 
   - Populates rows in Schedule sheet of Google sheet based on content from Rotation and Holiday sheets. Requires two dates be set in Rotation sheet (one for data and one for JC to initialize the iteration lower max date between the columns serves as the start marker for iteration). 

## <a name="futurefeats"></a> Future Features  

- **Dynamic/automated Schedule Updating and Absence Handling**:  
  Update schedules and trigger automatic cancellations for absences (e.g., Shawn's absences).  

- **Presentation Tracking**:  
  Count the number of presentations assigned to each person.  

- **Optimized Scheduling**:  
  Use a greedy algorithm to optimize the rotation based on presentation counts instead of a fixed rotation. Need to consider how much offset to give new lab members else they may be immediately next to present.  

