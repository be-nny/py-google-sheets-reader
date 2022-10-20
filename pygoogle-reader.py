"""
--- Library for reading google sheets ---

Provide the ID for the Google sheets doc
and the sheet number (indexing starts at 0).
Selecting cells syntax:
'!X# :Y#' >> where 'X' and 'Y' are the starting and ending
cell column names, and '#' is the starting row number
and the ending row number.

Built by Ben Abbott :)
"""

import os
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
          'https://www.googleapis.com/auth/drive']

COOL_DOWNS = [5, 10, 20, 40, 60]
CURRENT_COOL_DOWN_INDEX = 0

def getAuth():
    token_file = input("Google Credentials File Path >>")
    try:
        token_json_path = token_file + "token.json"
        credentials_json_path = token_file + "credentials.json"
        creds = None
        # tokens stores the users access and refresh tokens, if it isnt there, it creates one, if not i reads from it
        if os.path.exists(token_json_path):
            creds = Credentials.from_authorized_user_file(token_json_path, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                # if the tokens file doesn't exist, take the credentials.json to create it.
                flow = InstalledAppFlow.from_client_secrets_file(credentials_json_path, SCOPES)
                creds = flow.run_local_server(port=0)
                with open(token_json_path, 'w') as token:
                    token.write(creds.to_json())
        return creds
    except Exception:
        print("An error occurred when trying to access your tokens.json and credentials.json")
        print("Please see the Googledoc for more info: ")
        print("https://developers.google.com/sheets/api/quickstart/python")

def getData(sheet_id, selected_range, creds, sheet_num):
    """
    Method for getting data from selected google sheet with selected range
    :type sheet_id: str
    :type selected_range: str
    :type creds: str
    :type sheet_num: int

    :parameter sheet_id: the id of the sheet to be read
    :parameter selected_range: the range of cells to be read
    :parameter creds: json creds: credentials generated from the OAuth screen
    :parameter sheet_num: the index of the sheet to be read
    :return: array containing the values for each cell
    :raises: HttpError dealt with by throttling back on requests
    """

    try:
        # loading the sheets
        service = build('sheets', 'v4', credentials=creds)
        sheet_meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = sheet_meta.get('sheets', '')

        # getting specific sheet
        sheet = sheets[sheet_num]
        title = sheet.get("properties", {}).get("title", "Sheet1")
        selectedRange = title + selected_range

        # requesting the pricing sheet data
        request = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=selectedRange).execute()
        values = request.get('values', [])

        return values

    except HttpError:
        handle429()
        return getData(sheet_id, selected_range, creds, sheet_num)

def handle429():
    """
    Handles the 429 error automatically as throttling is used
    """

    global CURRENT_COOL_DOWN_INDEX

    current_cool_down_time = COOL_DOWNS[CURRENT_COOL_DOWN_INDEX]
    current_max_cool_down = time.time() + current_cool_down_time
    current_time = time.time()

    if current_max_cool_down - current_time < current_cool_down_time:
        if CURRENT_COOL_DOWN_INDEX == len(COOL_DOWNS) - 1:
            CURRENT_COOL_DOWN_INDEX = 0
        else:
            CURRENT_COOL_DOWN_INDEX += 1
    else:
        CURRENT_COOL_DOWN_INDEX -= 1

    print("Throttling back for " + str(COOL_DOWNS[CURRENT_COOL_DOWN_INDEX]) + "s")
    time.sleep(COOL_DOWNS[CURRENT_COOL_DOWN_INDEX])

def info():
    print(__doc__)

def getDataInfo():
    print(getData.__name__)
    print(getData.__doc__)

def handel429Info():
    print(handle429.__name__)
    print(handle429.__doc__)
