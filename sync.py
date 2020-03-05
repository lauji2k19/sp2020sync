from __future__ import print_function
import asyncio
import pickle
import os.path
import json
import os
import operator
import socket
import json
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SHEET_ID = "1QzDvzIwPzpbcthRg69ZdbWdkDXbd1lOI1mI6f5YFh8Q"
CREDS_PATH = "credentials/credentials.json"
SHEET_TOKEN_PATH = "credentials/sheettoken.pickle"

class SyncToSheets:
    def __init__(self):
        creds = None
        if os.path.exists(SHEET_TOKEN_PATH):
            with open(SHEET_TOKEN_PATH, 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDS_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(SHEET_TOKEN_PATH, 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)
        self.sheet = service.spreadsheets()
    
    def read_match_data(self, comp_code):
        path = os.getcwd() + "/compdata/"
        matches = []
        for file in os.listdir(path):
            file_path = path+file
            if os.path.isfile(file_path) and file.split("_")[0] == comp_code:
                f = open(file_path)
                matches.append(json.loads(f.read()))
        matches.sort(key=lambda f: f["matchNumber"])
        return matches

    def match_jsons_to_comp_json(self, comp_code):
        path = os.getcwd() + "/compdata/"
        competition_json = {}
        competition_json["competitionCode"] = comp_code
        competition_json["matches"] = []
        for file in os.listdir(path):
            file_path = path+file
            if os.path.isfile(file_path) and file.split("_")[0] == comp_code:
                f = open(file_path)
                competition_json["matches"].append(json.loads(f.read()))
        competition_json["matches"].sort(key=lambda f: f["matchNumber"])
        f = open(f"{path}{comp_code}.json", "w")
        f.write(json.dumps(competition_json, indent=4))
        f.close()
        return competition_json

    def push_competition_data(self, comp_code):
        competition = self.match_jsons_to_comp_json(comp_code)
        sheet_info = []
        for match in competition["matches"]:
            sheet_info.append([match["teamNumber"], match["matchNumber"], match["alliance"], match["initLine"], match["autoLower"], match["autoOuter"],
                            match["autoInner"], match["lower"], match["outer"], match["inner"], match["rotation"], match["position"], match["park"],
                            match["hang"], match["level"], match["disableTime"]])
        body = {
            "majorDimension": "ROWS",
            "values": sheet_info
        }
        response = self.sheet.values().update(spreadsheetId=SHEET_ID, range="A2:P", valueInputOption="USER_ENTERED", body=body).execute()
        return f"SUCCESS: {response}"        
    
    def push_match_data(self, comp_code):
        matches = self.read_match_data(comp_code)
        if len(matches) == 0:
            return f"FAIL: No competition found with code: {comp_code}"
        sheet_info = []
        for match in matches:
            sheet_info.append([match["teamNumber"], match["matchNumber"], match["alliance"], match["initLine"], match["autoLower"], match["autoOuter"],
                            match["autoInner"], match["lower"], match["outer"], match["inner"], match["rotation"], match["position"], match["park"],
                            match["hang"], match["level"], match["disableTime"]])
        body = {
            "majorDimension": "ROWS",
            "values": sheet_info
        }
        response = self.sheet.values().update(spreadsheetId=SHEET_ID, range="A2:P", valueInputOption="USER_ENTERED", body=body).execute()
        return f"SUCCESS: {response}"

if len(sys.argv) != 2:
    print("Invalid syntax. The script should be run like: python sync.py COMPCODE")
else:
    sync = SyncToSheets()
    #response = sync.push_match_data(sys.argv[1])
    response = sync.push_competition_data(sys.argv[1])
    print(response)