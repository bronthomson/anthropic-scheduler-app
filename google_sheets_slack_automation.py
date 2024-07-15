#!/usr/bin/env python3
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# Set up Google Sheets credentials and client
creds = Credentials.from_service_account_file('/Users/bron/Documents/Maramataka-slack-automation/brons-test-project-18f917064051.json')
sheets_service = build('sheets', 'v4', credentials=creds)

# Set up Slack client
slack_token = os.environ.get('SLACK_BOT_TOKEN')
if not slack_token:
    raise ValueError("SLACK_BOT_TOKEN environment variable is not set")
slack_client = WebClient(token=slack_token)

SHEET_ID = '1QNyDuoFid_Nf03ROd8oJdr-oD3tshIJjWTFWtdQjO50'
CHANNEL_ID = 'C070SDFNL3C'

def get_last_slack_message():
    try:
        result = slack_client.conversations_history(channel=CHANNEL_ID, limit=1)
        messages = result.get('messages', [])
        return messages[0]['text'] if messages else None
    except SlackApiError as e:
        print(f"Error fetching last message: {e}")
        return None

import unicodedata

def strip_accents(text):
    return ''.join(char for char in
                   unicodedata.normalize('NFKD', text)
                   if unicodedata.category(char) != 'Mn')

def get_data_from_sheets(lookup_value):
    try:
        range_name = 'Sheet1!A1:B'  # Adjust if your sheet has a different name
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
                    return row[1] if len(row) > 1 else None
        
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
    last_message = get_last_slack_message()
    print(f"Last message: {last_message}")
    if last_message:
        print(f"Searching for: '{last_message}'")
        data = get_data_from_sheets(last_message)
        if data:
            send_message_to_slack(f"{data}")
        else:
            send_message_to_slack(f"Sorry, I couldn't find any data related to '{last_message}'.")
    else:
        send_message_to_slack("No previous messages found to determine context.")

if __name__ == '__main__':
    main()
