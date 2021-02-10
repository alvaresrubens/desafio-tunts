import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1xBoRULQD8vNtOTG7K78FRM9KyMj7IldxxUQWN9u7GzM'
RANGE_NAME = 'engenharia_de_software!A4:H'


def google_connection():
    """
    Checks for a valid token and creates one when none is found.
    """
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
    return service.spreadsheets()


def set_google_data(results):
    """
    Receives classified grades and updated the google sheet accordingly.
    """
            
    print('Step 3 - Updating Google Sheets')
    
    sheet = google_connection()
   
    cells_range = 'engenharia_de_software!G4:H'  
    body = {
        'values': results
    }
    result = sheet.values().update(spreadsheetId=SPREADSHEET_ID, valueInputOption='RAW', range=cells_range, body=body).execute()
    print('Finished! {0} cells updated.'.format(result.get('updatedCells')))


def get_google_data():
    """
    Scrapes google sheets for student grades of a given range.
    """
    print('Step 1 - Downloading data from Google Sheets...')
    sheet = google_connection()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    
    return result.get('values', [])


def classify_grades(downloaded_data):
    """
    Receives raw grades and return a array with student grade classification and the necessary final exam grade when relevant.
    """
    print('Step 2 - Classifying students grades')
    grades = []
    for row in downloaded_data:
        abscences = int(row[2])
        mean = (int(row[3]) +int(row[4])+int(row[5]))/30

        if (abscences > 15 ):
            grades.append(['Reprovado por falta', 0])
        elif ( mean < 5 ):

            grades.append(['Reprovado por nota', 0])
        elif ( 5 <= mean < 7 ):
            naf = 10 - mean
            if (naf <= 5 ):
                naf = round (naf + 0.5)
                grades.append(['Exame final', naf])
        else:
            grades.append(['Aprovado', 0])
    return grades    


def main():
    
    downloaded_data = get_google_data()
   
    if not downloaded_data:
        print('No data found.')
    else:
        grades = classify_grades(downloaded_data)
        set_google_data(grades)
        

if __name__ == '__main__':
    print('Starting execution')
    main()