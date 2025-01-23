import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from datetime import datetime

TOKEN_FILE = "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/gmail.send"
]

def refresh_token():
    # Load credentials from the token file
    print(TOKEN_FILE)
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    else:
        print("No token file found. Run get_token() first.")
        return

    # Refresh the token if a refresh token is available
    if creds and creds.refresh_token:
        creds.refresh(Request())
        # Save the refreshed token back to the file
        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
        print(f"Token refreshed at {datetime.now()}")
    else:
        print("No refresh token available. Run get_token() again.")

if __name__ == "__main__":
    refresh_token()
