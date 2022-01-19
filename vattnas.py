from __future__ import print_function
import datetime
import time
import json
from os import name
import os.path
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar',
          'https://www.googleapis.com/auth/gmail.readonly']

month = {
    'jan': '01',
    'feb': '02',
    'mar': '03',
    'apr': '04',
    'maj': '05',
    'jun': '06',
    'jul': '07',
    'aug': '08',
    'sep': '09',
    'okt': '10',
    'nov': '11',
    'dec': '12',
}


def booking_name(body):

    split_line = []
    first_name = []
    last_name = []
    for line in body.splitlines():
        split_line = line.split()
        if split_line:
            if split_line[0] == 'Förnamn:':
                first_name = list(split_line[1:])
            elif split_line[0] == 'Efternamn:':
                last_name = list(split_line[1:])
    name_list = list(first_name)
    for ln in last_name:
        name_list.append(ln)
    return " ".join(name_list)


def booking_dates(body):
    if 'Datum:' not in body:
        return None, None

    for line in body.splitlines():
        split_line = line.split()
        if split_line and split_line[0] == 'Datum:':
            start = f"{split_line[5]}-{month[split_line[4]]}-{split_line[3]}"
            end = f"{split_line[9]}-{month[split_line[8]]}-{split_line[7]}"
            return start, end
    return None, None


def booking_id(body):

    split_line = []
    for line in body.splitlines():
        split_line = line.split()
        if split_line and split_line[0] == 'Bokningsnummer:':
            return split_line[1]
    return "0"


def setup_creds():
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def scan_and_get_message_bodies(service_gmail):
    print("Fetching all messages matching the query...")
    results = service_gmail.users().messages().list(
        userId='me', q='from:noreply@citybreak.com').execute()
    messages = results.get('messages', [])
    message_bodies = []

    if not messages:
        print('No messages found.')
    else:
        print(f"{len(messages)} messages found. Looking for bookings...")
        for message in messages:
            msg = service_gmail.users().messages().get(
                userId='me', id=message['id']).execute()
            for header in msg['payload']['headers']:
                if (
                    header['name'] == 'Subject'
                    and 'NY BOKNING' in header['value']
                ):
                    body = base64.urlsafe_b64decode(msg['payload'].get(
                        "body").get("data").encode("ASCII")).decode("utf-8")
                    message_bodies.append(body)
    print(f"{len(message_bodies)} bookings found.")
    return message_bodies


def update_db(service, message_bodies, events):

    print("Opening json file...")
    with open('events.json') as events_file:
        json_str = events_file.read()
        if json_str:
            events = json.loads(json_str)

    event_ids = [booking_id(event["description"])
                 for event in events['events']] if events else []

    print("Adding events to json if not already in json:")
    for body in message_bodies:
        start_date, end_date = booking_dates(body)
        desc = "".join(line + "\n" for line in body.splitlines())
        summary = booking_name(body)
        id = booking_id(body)

        event = {
            'summary': summary,
            'description': desc,
            'start': {'date': start_date},
            'end': {'date': end_date}
        }

        if event['summary'] and id not in event_ids:
            print("a", end="")
            events['events'].append(event)
            with open('events.json', 'w') as events_file:
                events_file.write(json.dumps(events))
        else:
            print(".", end="")
    print("Done")
    return events


def update_calendar(service, events):
    print('Fetching all calendar events...')
    cal_events = service.events().list(calendarId='primary').execute()
    print(f"{len(cal_events['items'])} calendar events found.")
    exists = False
    print('Adding events in json to calendar if not already in calendar:')
    for event in events['events']:
        for cal_event in cal_events['items']:
            if event['summary'] == cal_event['summary']:
                exists = True
        if not exists:
            print("a")
            e = service.events().insert(calendarId='primary', body=event).execute()

            print(f'''*** {e['summary'].encode('utf-8')} event added:
            Start: {e['start']['date']}
            End:   {e['end']['date']}''')
        else:
            print(".", end="")
        exists = False
    print('Done')


def main():

    creds = setup_creds()

    service_cal = build('calendar', 'v3', credentials=creds)
    service_gmail = build('gmail', 'v1', credentials=creds)

    events = {
        'events': []
    }

    message_bodies = scan_and_get_message_bodies(service_gmail)
    events = update_db(service_cal, message_bodies, events)
    update_calendar(service_cal, events)


if __name__ == '__main__':
    main()
