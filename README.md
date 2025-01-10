# auto-lab-reminders

This Python application allows users to set schedules in Google Sheets and send Google Calendar invites (emails) and MS Teams reminders based on the schedule. Accessing calendar data is done through Google API. 

## Requirements  
- **Python Version**: Python >= 3.8  
  - *Recommendation*: If you don't already have a valid version of Python, consider installing the latest version.  

## Python Dependencies  
To install the required Python libraries, use the following command:  
```bash
pip install argparse pandas google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread pymsteams
```

## Components  

The application consists of the following:  

### Google Sheet  

The main Google Sheet contains the following sheets:  

1. **Rotation**  
   Specifies the rotation logic for scheduling events or tasks.  

2. **Emails**  
   Handles sending email invites to participants based on the schedule.  

3. **Holidays**  
   - Default holidays are hardcoded based on BCM holidays (falling on Thursdays) and include a 2-week winter break (which don't need to appear on this sheet).  
   - This sheet is to primarily to add holidays or mark Shawn's absences that are not BCM holidays  

4. **Schedule**  
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
- `zoomextras`: additional text appended onto calendar invite message/description 

[teams]
- `maxevents`: number of events to show on MS teams notification

### Scripts  

The application includes two Python scripts for execution, which can be run manually or scheduled via a cron job:  

1. **`cal_invite.py`**  
   - Handles sending calendar invites based on the schedule. This will send a email for instances where lab meeting has been canceled due to Shawn's absence or other circumstances.
   - Manual runs of the script get the next proximal event, while including `--auto` flag in a cron job triggers invites only for events a week away (hardcoded).

2. **`msteams_remind.py`**  
   - Sends notifications via Microsoft Teams based on the schedule.  

3. (optional) **`generate_schedule.py`** 
   - Populates rows in Schedule sheet of Google sheet based on content from Rotation and Holiday sheets. 

## Future Features  

- **Dynamic/automated Schedule Updating and Absence Handling**:  
  Update schedules and trigger automatic cancellations for absences (e.g., Shawn's absences).  

- **Presentation Tracking**:  
  Count the number of presentations assigned to each person.  

- **Optimized Scheduling**:  
  Use a greedy algorithm to optimize the rotation based on presentation counts instead of a fixed rotation. Need to consider how much offset to give new lab members else they may be immediately next to present.  

