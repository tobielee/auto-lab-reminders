import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
import configparser

# Load configuration
config = configparser.ConfigParser()
config.read('../cal_config.cfg')
labmeeting_settings = config['labmeeting']

# File paths for credentials and token
CREDENTIALS_FILE = f"../{labmeeting_settings['usercreds']}"  # Replace with the path to your downloaded client credentials
TOKEN_FILE = "token.json"  # This will be created automatically

# Scopes you want to enable
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/gmail.send"
]

def get_token():
    creds = None

    # Check if the token already exists
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # If there are no valid credentials, prompt the user to log in
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0, access_type='offline', prompt='consent')

        # Save the token for future use
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    print("Token saved to", TOKEN_FILE)
    return creds

if __name__ == "__main__":
    get_token()

