import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main():
    creds = None

    if os.path.exists("secret/token.json"):
        creds = Credentials.from_authorized_user_file("secret/token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
            "secret/credentials.json", SCOPES
         )
        creds = flow.run_local_server(port=0)

    with open("secret/token.json", "w") as token:
      token.write(creds.to_json())

    try:

        service = build("calendar", "v3", credentials=creds)

        event = {
            'summary': 'Event NEW',
            'location': 'Nishi shinjuku, japan',
            'description': 'this is description',
            'colorId': 3,
            'start': {
                'dateTime': '2024-03-25T09:00:00-07:00',
                'timeZone': 'America/Los_Angeles',
            },
            'end': {
                'dateTime': '2024-03-25T09:20:00-07:00',
                'timeZone': 'America/Los_Angeles',
            },
            'recurrence': [
                'RRULE:FREQ=DAILY;COUNT=1'
            ],
            'attendees': [
                {'email': 'anmolgera01@gmail.com'},
            ]
        }

        event = service.events().insert(calendarId='primary', body=event).execute()
        print(f"Event created: {event.get('htmlLink')}")

    except HttpError as error:
        print(f"An error occurred: {error}")



if __name__ == "__main__":
  main()