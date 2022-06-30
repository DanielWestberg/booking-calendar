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

_month = {
    'jan': '01',
    'feb': '02',
    'mar': '03',
    'apr': '04',
    'maj': '05',
    'may': '05',
    'jun': '06',
    'jul': '07',
    'aug': '08',
    'sep': '09',
    'okt': '10',
    'oct': '10',
    'nov': '11',
    'dec': '12',
    'JAN': '01',
    'FEB': '02',
    'MAR': '03',
    'APR': '04',
    'MAJ': '05',
    'MAY': '05',
    'JUN': '06',
    'JUL': '07',
    'AUG': '08',
    'SEP': '09',
    'OKT': '10',
    'OCT': '10',
    'NOV': '11',
    'DEC': '12',
}


def get_booking_name(body):
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


def get_start_and_end_dates(body):
    if 'Datum:' not in body:
        return None, None

    for line in body.splitlines():
        split_line = line.split()
        if split_line and split_line[0] == 'Datum:':
            start = f"{split_line[5]}-{_month[split_line[4]]}-{split_line[3]}"
            end = f"{split_line[9]}-{_month[split_line[8]]}-{split_line[7]}"
            return start, end
    return None, None


def get_booking_id(body):
    split_line = []
    for line in body.splitlines():
        split_line = line.split()
        if split_line and split_line[0] == 'Bokningsnummer:':
            return split_line[1]
    return "0"


def get_booking_date(body):
    split_line = []
    for line in body.splitlines():
        split_line = line.split()
        if split_line and split_line[0] == 'Bokningsdatum:':
            return f"20{split_line[1][5] + split_line[1][6]}-{_month[split_line[1][2] + split_line[1][3] + split_line[1][4]]}-{split_line[1][0] + split_line[1][1]}"
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
    message_bodies_new = []
    message_bodies_changed = []
    message_bodies_canceled = []

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
                    and ('NY BOKNING') in header['value']
                ):
                    body = base64.urlsafe_b64decode(msg['payload'].get(
                        "body").get("data").encode("ASCII")).decode("utf-8")
                    message_bodies_new.append(body)
                elif (
                    header['name'] == 'Subject'
                    and ('ÄNDRAD BOKNING') in header['value']
                ):
                    body = base64.urlsafe_b64decode(msg['payload'].get(
                        "body").get("data").encode("ASCII")).decode("utf-8")
                    message_bodies_changed.append(body)
                elif (
                    header['name'] == 'Subject'
                    and ('AVBOKAD BOKNING') in header['value']
                ):
                    body = base64.urlsafe_b64decode(msg['payload'].get(
                        "body").get("data").encode("ASCII")).decode("utf-8")
                    message_bodies_canceled.append(body)

    print(f"{len(message_bodies_new)} new bookings found.")
    print(f"{len(message_bodies_changed)} changed bookings found.")
    print(f"{len(message_bodies_canceled)} canceled bookings found.")
    return message_bodies_new, message_bodies_changed, message_bodies_canceled


def add_or_change_events(message_bodies, events):
    for body in message_bodies:
        start_date, end_date = get_start_and_end_dates(body)
        desc = "".join(line + "\n" for line in body.splitlines())
        summary = get_booking_name(body)
        booking_id = get_booking_id(body)
        booking_date = get_booking_date(body)

        event = {
            'booking_id': booking_id,
            'booking_date': booking_date,
            'start': {'date': start_date},
            'end': {'date': end_date},
            'summary': f"VATTNÄS: {summary}, {booking_id}",
            'description': desc
        }
        if booking_id in events:
            if event['booking_date'] > events[booking_id]['booking_date']:
                events[booking_id] = event
                print("c", end="")
            else:
                print(".", end="")
        else:
            events[booking_id] = event
            print("a", end="")
    print()
    return events


def cancel_event(message_bodies_canceled, events):
    for body in message_bodies_canceled:
        booking_id = get_booking_id(body)
        if events.pop(booking_id):
            print("d", end="")
        else:
            print(".", end="")
    print()
    return events


def update_db(message_bodies_new, message_bodies_changed, message_bodies_canceled):
    print("Opening json file...")
    with open('events.json') as events_file:
        if json_str := events_file.read():
            events = json.loads(json_str)
    print("Adding new events to json if not already in json...")
    events = add_or_change_events(message_bodies_new, events)
    print("Replacing events that were changed...")
    events = add_or_change_events(message_bodies_changed, events)
    print("Removing events that were canceled...")
    events = cancel_event(message_bodies_canceled, events)
    print("Updating json...")
    with open('events.json', 'w') as events_file:
        events_file.write(json.dumps(events))
    print("Done")
    return events


def get_booking_events(cal_events_all):
    cal_events = {}
    for cal_event in cal_events_all['items']:
        if "BOKNING" in cal_event['description'] and "VISIT DALARNA" in cal_event['description']:
            cal_events[get_booking_id(cal_event['description'])] = cal_event
    return cal_events


def update_calendar(service, events):
    print('Fetching all calendar events...')
    cal_events_all = service.events().list(calendarId='primary').execute()
    cal_events = get_booking_events(cal_events_all)
    print(
        f"{len(cal_events_all['items'])} calendar events found. Looking for bookings...")
    print(f"{len(cal_events)} booking calendar events found.")
    exists = False
    changed = False
    print('Updating calendar according to json...')
    for event in events.values():
        booking_id = event['booking_id']
        if booking_id in cal_events:
            exists = True
            if event['description'] != cal_events[booking_id]['description']:
                service.events().delete(calendarId='primary',
                                        eventId=cal_events[booking_id]['id']).execute()
                changed = True
        if not exists:
            print("a", end="")
            e = service.events().insert(calendarId='primary', body=event).execute()

            print(f'''*** {e['summary'].encode('utf-8')} event added:
            Start: {e['start']['date']}
            End:   {e['end']['date']}''')
        elif changed:
            print("c", end="")
            e = service.events().insert(calendarId='primary', body=event).execute()

            print(f'''*** {e['summary'].encode('utf-8')} event updated:
            Start: {e['start']['date']}
            End:   {e['end']['date']}''')
        else:
            print(".", end="")
        exists = False
        changed = False
    print('\nRemoving canceled events from calendar...')
    for booking_id in cal_events.keys():
        if booking_id not in events:
            print("d")
            service.events().delete(calendarId='primary',
                                    eventId=cal_events[booking_id]['id']).execute()
            print(f"*** {cal_events[booking_id]['summary']} event removed")
        else:
            print(".", end="")
    print('Done')


def main():
    creds = setup_creds()

    service_cal = build('calendar', 'v3', credentials=creds)
    service_gmail = build('gmail', 'v1', credentials=creds)

    message_bodies_new, message_bodies_changed, message_bodies_canceled = scan_and_get_message_bodies(
        service_gmail)
    events = update_db(message_bodies_new, message_bodies_changed,
                       message_bodies_canceled)
    update_calendar(service_cal, events)


if __name__ == '__main__':
    main()
