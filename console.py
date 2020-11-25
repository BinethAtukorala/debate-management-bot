from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sqlite3
from sqlite3 import Error

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print(sqlite3.version)
        return conn
    except Error as e:
        print(e)
    

def create_table(conn, create_table_sql):
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)

def new_debate(conn, date, venue, time):
    try:
        c = conn.cursor()
        command = " INSERT INTO debates (date, time, venue) VALUES ('" + str(date) +"' , '" + str(time) + "' , '" + str(venue) + "');"
        c.execute(command)
    except Error as e:
        print(e)

def no_of_participants(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range="J5:L5").execute()
    rows = result.get('values', [])
    return_list = [int(rows[0][0]), int(rows[0][2])]
    return return_list

conn = create_connection(r"pythonsqlite.db")

# Create table command

sql_create_debates_table = """ CREATE TABLE IF NOT EXISTS debates (
                                    id INTEGER PRIMARY KEY,
                                    time VARCHAR NOT NULL,
                                    date VARCHAR NOT NULL,
                                    venue VARCHAR NOT NULL,
                                    sheet VARCHAR NOT NULL
                                );
                            """

if(conn != None):
    create_table(conn, sql_create_debates_table)

else:
    print("Err, no db connection")


# Google sheets setup

creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

spreadsheet_id = "1S1xLC0K_PzM727MKDYXpfe9u381Sw49XFuhYNjKLcvU"

# Main loop

while(True):

    cmd = input("Enter command (new, status, end, participate, assign, finalize) : ")

    # Commands

    if(cmd == "new"):

        if(spreadsheet_id!=""):
            print("You have to end the current session before starting a new one...")
            continue

        date = input("Enter date: ")
        time = input("Enter time: ")
        venue = input("Enter venue: ")

        spreadsheet = {
            'properties': {
                'title': "MUN Test"
            }
        }

        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields="spreadsheetId").execute()
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        print("Spreadsheet ID: {0}".format(spreadsheet_id))
        print("Link: https://docs.google.com/spreadsheets/d/{0}/edit#gid=0".format(spreadsheet_id))

        values = [
            [
                "Name", "Topic", "Committee", "Country", "Attendance", "Sent", "", "Country Pool", "", "Time", "Date", "Venue", "No of participants (lim: 200)"
            ],
            [
                "", "", "", "", "", "", "", "", "", time, date, venue
            ], [], 
            [
                "", "", "", "", "", "", "", "", "", "No of participants", "", "No of countries entered"
            ],
            [
                "", "", "", "", "", "", "", "", "", "=COUNTA(A2:A202)", "", "=COUNTA(H2:H202)"
            ]
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range="A1:L5",
            valueInputOption="USER_ENTERED", body=body).execute()
        
        print('{0} cells updated.'.format(result.get('updatedCells')))


    elif(cmd == "status"):
        print('{0} participants registered.'.format(no_of_participants(service)))
    
    elif(cmd == "end"):
        spreadsheet_id = ""
    
    elif(cmd == "participate"):
        name = input("Enter participant name : ")

        body = {
            'values': [
                [
                    name, "", "", "", "false", "false"
                ]
            ]
        }

        row_no = no_of_participants(service)[0] + 2

        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range="A{0}:F{0}".format(row_no),
            valueInputOption="USER_ENTERED", body=body).execute()
        
        print('{0} cells updated.'.format(result.get('updatedCells')))

    elif("assign"): 
        no_of_countries = no_of_participants(service)[1]

        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range="H2:H" + str(no_of_countries + 1)).execute()
        rows = result.get('values', [])
        countries = []
        for x in rows:
            countries.append(x[0])
        print(countries)