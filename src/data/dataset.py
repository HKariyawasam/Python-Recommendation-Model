import pandas as pd
from openpyxl import load_workbook

from google.oauth2 import service_account
from googleapiclient.discovery import build


# Set up credentials
credentials = service_account.Credentials.from_service_account_file('credentials.json')
drive_service = build('drive', 'v3', credentials=credentials)

# Specify the file ID of the Excel file on Google Drive
file_id = '18lfP1BZOfpy1S1Rjgrr1G7sy0JiknDQV'  # Replace with the actual file ID
file_id2 = '1xjEEOrjQkUBLDtb1UHGrurKkVdmp0meS'

# Download the Excel file and read it using pandas
request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
response = request.execute()
with open('file.xlsx', 'wb') as file:
    file.write(response)

request2 = drive_service.files().export_media(fileId=file_id2, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
response2 = request2.execute()
with open('file2.xlsx', 'wb') as file:
    file.write(response)



def readActivityDataFrame():
    activities_df = pd.read_excel('file.xlsx')
    return activities_df

def readUsersCompletionDataFrame():
    completerates_df = pd.read_excel('file2.xlsx')
    return completerates_df

def updateUserCompletionRates(data):
    userId = int(data['userId'])
    activityId = int(data['activityId'])
    complete_score = int(data['complete_score'])
    satisfaction_score = float(data['satisfaction_score'])

    wb_append = load_workbook("data/completeratecheckappend.xlsx")

    sheet = wb_append.active
    row = (userId,activityId, complete_score,satisfaction_score)
    sheet.append(row)
    wb_append.save('data/completeratecheckappend.xlsx')
