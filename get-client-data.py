from __future__ import print_function
import pickle
import os.path
import sqlite3
import json
import pprint
import re
import time
from copy import deepcopy
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

spreadsheets = [
    'timesheet ID', # This is a timesheet other data is compared to
    'google sheet IDs also go in this list'
]
attendanceRanges = [
    "name of attendance sheet and range" # for example: JAN'20_GR!A2:Y52
]
dataRanges = [
    # below are the examples of sheet names and their ranges
    'Jan!A1:AF5',
    'Feb!A1:AF5',
    'Mar!A1:AF5',
    'Apr!A1:AF5',
    'May!A1:AF5',
    'June!A1:AF5',
    'July!A1:AF5',
    'Aug!A1:AF5',
    'Sep!A1:AF5',
    'Oct!A1:AF5',
    'Nov!A1:AF5',
    'Dec!A1:AF5'
]

sheetObj = []

def main(sheetArr, rangeArr):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    sheetData = service.spreadsheets().values().get(spreadsheetId=sheetArr,
                                range=rangeArr).execute()
    rowData = sheetData.get('values', [])

    sheetObject = service.spreadsheets().get(spreadsheetId=sheetArr).execute()
    print('making a google sheets request')


    if not rowData:
        print('No data found.')
    else:
        return [sheetObject, rowData]

if __name__ == '__main__':
    attenGreen = main(spreadsheets[0], attendanceRanges[0])
    attenBlue = main(spreadsheets[0], attendanceRanges[1])
    atten = deepcopy(attenBlue[1] + attenGreen[1])
    weekdays = atten[0][2:]
    ispObj = {
        'weekdays': weekdays
        }
    for i in range(1, len(spreadsheets)):
        j = 1
        s = main(spreadsheets[i], dataRanges[0])
        pattern = atten[j][0].lower()
        searchIn = re.compile(r"("+s[0]['properties']['title'].lower()[0:len(atten[j][0])]+")")
        # below, an excpetion is written if a client's name is used inconsistently in different sheets
        if s[0]['properties']['title'].lower()[0:len(atten[j][0])] in "John M. Smith":
            searchIn = re.compile(r"smith, john")

        while atten[j][0] == "":
            j += 1

        while not searchIn.search(pattern):
            j += 1
            try:
                print("notmatched ", j-1, atten[j-1][0].lower(), " and ", i, s[0]['properties']['title'].lower())
                pattern = atten[j][0].lower()
                searchIn = re.compile(r"("+s[0]['properties']['title'].lower()[0:len(atten[j][0])]+")")
                if s[0]['properties']['title'].lower()[0:len(atten[j][0])] in "John M. Smith":
                    searchIn = re.compile(r"smith, john")
            except IndexError:
                print(s[0]['properties']['title'], " not found in data set")
                break
            while atten[j][0] == "":
                j += 1
                try:
                    pattern = atten[j][0].lower()
                    searchIn = re.compile(r"("+s[0]['properties']['title'].lower()[0:len(atten[j][0])]+")")
                    if s[0]['properties']['title'].lower()[0:len(atten[j][0])] in "John M. Smith":
                        searchIn = re.compile(r"smith, john")
                except IndexError:
                    print(s[0]['properties']['title'], " not found in data set")
                    break

        if searchIn.search(pattern):
            try:
                print('item', "i", i, "j", j)
                print(s[0]['properties']['title'].lower())
                ispObj.setdefault(s[0]['properties']['title'].lower(), {
                    'data-set': s[1],
                    'attendance': atten[j]
                    })
                print('----END MATCH----')
                for i in range(0, 2):
                    time.sleep(7)
                    print('pausing for Google API limit ', i+1, "of 2")
                
            except IndexError:
                print("Final check not found in data set")
                pprint.pprint(atten)
    
    cliJSON = open('client-data.json', 'w')
    cliJSON.write(json.dumps(ispObj))
