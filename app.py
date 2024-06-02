#imports 

from flask import Flask, render_template, request, redirect
import os.path
import pickle
import google.auth.transport.requests
import google_auth_oauthlib.flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)


# delete auto generated token.pickle (after running script) file if scopes change
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file']


# range of google spreadsheet
RANGE_NAME = 'Sheet1!A1:E1'


def get_credentials():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                # credential file from google cloud
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=5000)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def create_spreadsheet(service):
    spreadsheet_body = {
        'properties': {
            'title': 'Job Applications Tracker'
        },
        'sheets': [
            {
                'properties': {
                    'title': 'Sheet1'
                }
            }
        ]
    }
    request = service.spreadsheets().create(body=spreadsheet_body)
    response = request.execute()
    return response['spreadsheetId']

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/submit', methods=['POST'])
def submit():
    job_name = request.form['jobName']
    job_link = request.form['jobLink']
    job_description = request.form['jobDescription']
    apply_date = request.form['applyDate']
    job_status = request.form['jobStatus']

    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # get spreadsheet ID from file
    spreadsheet_id_file = 'spreadsheet_id.txt'
    if os.path.exists(spreadsheet_id_file):
        with open(spreadsheet_id_file, 'r') as f:
            spreadsheet_id = f.read().strip()
    else:
        # new spreadsheet if no ID 
        try:
            spreadsheet_id = create_spreadsheet(service)
            with open(spreadsheet_id_file, 'w') as f:
                f.write(spreadsheet_id)
        except HttpError as error:
            print(f"An error occurred: {error}")
            return redirect('/')

    sheet = service.spreadsheets()

    values = [
        [job_name, job_link, job_description, apply_date, job_status]
    ]
    body = {
        'values': values
    }
    result = sheet.values().append(
        spreadsheetId=spreadsheet_id, range=RANGE_NAME,
        valueInputOption="RAW", body=body).execute()

    row = result['updates']['updatedRange'].split('!')[1].split(':')[0][1:]

    # conditional formatting
    apply_conditional_formatting(service, spreadsheet_id)

    return redirect('/')

def apply_conditional_formatting(service, spreadsheet_id):
    
    # sheet id
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']

    # updating cell colors of job status in sheets
    color_map = {
        'applied': {'red': 1.0, 'green': 1.0, 'blue': 0.0},  #yellow
        'interview': {'red': 0.0, 'green': 0.0, 'blue': 1.0},  #blue
        'offered': {'red': 0.0, 'green': 1.0, 'blue': 0.0},  #green
        'rejected': {'red': 1.0, 'green': 0.0, 'blue': 0.0}  #red
    }

    # empty 
    requests = []

    for status, color in color_map.items():
        requests.append({
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': [{
                        'sheetId': sheet_id,
                        #start from second row
                        'startRowIndex': 1,
                        #column E is index 4
                        'startColumnIndex': 4,
                        'endColumnIndex': 5
                    }],
                    'booleanRule': {
                        'condition': {
                            'type': 'TEXT_EQ',
                            'values': [{
                                'userEnteredValue': status
                            }]
                        },
                        'format': {
                            'backgroundColor': color
                        }
                    }
                },
                'index': 0
            }
        })

    # formatting
    batch_update_request_body = {
        'requests': requests
    }

    response = service.spreadsheets().batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = batch_update_request_body
    ).execute()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
