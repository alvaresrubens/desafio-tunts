from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1xBoRULQD8vNtOTG7K78FRM9KyMj7IldxxUQWN9u7GzM'
SAMPLE_RANGE_NAME = 'engenharia_de_software!A4:H'


def set_data_google(results):

    sheet = google_connection()
   
    range = 'engenharia_de_software!G4:H'  
 
    values = results
    body = {
        'values': values
    }
    result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, valueInputOption='RAW', range=range, body=body).execute()
    print('Finished! {0} cells updated.'.format(result.get('updatedCells')))

def get_google_data(): 
    
    sheet = google_connection()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    return values

def google_connection():
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
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    return sheet
def main():
    
    print('Step 1 - Downloading data from Google Sheets...')
    downloaded_data = get_google_data()
    
   
    if not downloaded_data:
        print('No data found.')
    else:
        print('Step 2 - Classifying students grades')

        grades = []
        for row in downloaded_data:
            
            abscences = int (row[2])
            mean = (int(row[3]) +int(row[4])+int(row[5]))/30

            if (abscences > 15 ):
                grades.append(['Reprovado por falta', 0])
                
            elif ( mean < 5 ):
                grades.append(['Reprovado por nota', 0])
            elif ( 5 <= mean < 7 ):
                naf = 10 - mean
                if (naf < 5 ):
                    naf = round (naf + 0.5)
                grades.append(['Exame final', naf])

            else:
                grades.append(['Aprovado', 0])
           
        print('Step 3 - Updating Google Sheets')

        set_data_google(grades)


if __name__ == '__main__':
    print('Starting execution')
    main()