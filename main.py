#!/usr/bin/env python3
import os
import json
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
import unicodedata

# Set up Google credentials and clients
creds_json = os.environ.get('GOOGLE_CREDENTIALS')
creds_dict = json.loads(creds_json)
creds = Credentials.from_service_account_info(creds_dict)
sheets_service = build('sheets', 'v4', credentials=creds)
calendar_service = build('calendar', 'v3', credentials=creds)

# Set up Slack client
slack_token = os.environ.get('SLACK_BOT_TOKEN')
if not slack_token:
    raise ValueError("SLACK_BOT_TOKEN environment variable is not set")
slack_client = WebClient(token=slack_token)

SHEET_ID = os.environ.get('SHEET_ID')
CHANNEL_ID = os.environ.get('CHANNEL_ID')
CALENDAR_ID = os.environ.get('CALENDAR_ID')

def get_todays_event():
    today = datetime.utcnow().date()
    tomorrow = today + timedelta(days=1)
    
    events_result = calendar_service.events().list(
        calendarId=CALENDAR_ID,
        timeMin=today.isoformat() + 'T00:00:00Z',
        timeMax=tomorrow.isoformat() + 'T00:00:00Z',
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    
    events = events_result.get('items', [])
    
    if events:
        return events[0]['summary']
    else:
        return None

def strip_accents(text):
    return ''.join(char for char in
                   unicodedata.normalize('NFKD', text)
                   if unicodedata.category(char) != 'Mn')

def get_data_from_sheets(lookup_value):
    try:
        range_name = 'Sheet1!A:B'  # Fetch both columns A and B
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=range_name).execute()
        values = result.get('values', [])
        
        if not values:
            print('No data found in the sheet.')
            return None
        
        normalized_lookup = strip_accents(lookup_value.lower())
        print(f"Normalized lookup value: {normalized_lookup}")
        
        for row in values:
            if row:  # Check if row is not empty
                sheet_value = row[0]
                normalized_sheet_value = strip_accents(sheet_value.lower())
                print(f"Comparing: '{normalized_lookup}' with '{normalized_sheet_value}'")
                if normalized_lookup == normalized_sheet_value:
                    return row  # Return the entire row (both columns)
        
        print(f"No match found for '{lookup_value}'")
        return None
    except Exception as e:
        print(f"Error fetching data from Sheets: {e}")
        return None

def send_message_to_slack(message):
    try:
        response = slack_client.chat_postMessage(channel=CHANNEL_ID, text=message)
        print(f"Message sent: {response['ts']}")
    except SlackApiError as e:
        print(f"Error sending message: {e}")

def main():
    event = get_todays_event()
    print(f"Today is: {event}")
    if event:
        print(f"Searching for: '{event}'")
        data = get_data_from_sheets(event)
        if data:
            send_message_to_slack(f"Today is {data[0]}\n {data[1]}")
        else:
            send_message_to_slack(f"Sorry, I couldn't find any data related to today's event: '{event}'.")
    else:
        send_message_to_slack("No event found for today.")

if __name__ == '__main__':
    main()
