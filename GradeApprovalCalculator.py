import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

def average(lst):
    size = len(lst)
    lst = map(int, lst)
    return sum(lst) / size
    
def update(status, exam, range):
    body = {"values": [[status, exam]]}
    result = (
    service.spreadsheets()
    .values()
    .update(
        spreadsheetId=SAMPLE_SPREADSHEET_ID,
        range=range,
        valueInputOption='RAW',
        body=body,
    )
    .execute()
)

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# the ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1RXM7PvfGxLBOxhanhrZtlv4i1Y5qp7fqAJ7InrVjqrY'

creds = None

if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# if there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        
service = build('sheets', 'v4', credentials=creds)  

# call the Sheets API
sheet = service.spreadsheets()

# parse sheet infos by rows
sheet = service.spreadsheets().values().batchGet(
         spreadsheetId=SAMPLE_SPREADSHEET_ID,
         ranges='engenharia_de_software',
         majorDimension='ROWS').execute()

# split into rows values
value_range = sheet['valueRanges'][0]['values']

# extract classes number
tt_class = int(value_range[1][0].split(':')[1])

# iterate over list to get average and frequency
for i in range(len(value_range)):
    if value_range[i][0].isnumeric():
        
        # calculate average
        grad_average = round(average(value_range[i][3:6]))/10
        
        # calculate frequency
        freq = 100*(tt_class - int(value_range[i][2]))/tt_class
        
        # sheet range to update
        loc = 'G' + str(i+1) + ':H' + str(i+1)
        
        a = 'J' + str(i+1) + ':K' + str(i+1)
        result = update(grad_average,grad_average, a)
        
        if freq < 75:
            exam = 0
            result = update('Reprovado por Falta', exam, loc)
        elif grad_average >= 7:
            exam = 0
            result = update('Aprovado', exam, loc)
        elif grad_average < 5:
            exam = 0
            result = update('Reprovado por Nota', exam, loc)
        else:
            exam = 10 - grad_average
            result = update('Exame Final', exam, loc)
            
            
            
        

