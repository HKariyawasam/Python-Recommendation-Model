import pandas as pd
from flask import jsonify
from openpyxl import load_workbook


import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define the scope and credentials for accessing the Google Drive and Sheets APIs
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)


# Authenticate and authorize the credentials
client = gspread.authorize(credentials)


spreadsheet_url1 = 'https://docs.google.com/spreadsheets/d/10AQCTZt-mpp0xQ5T4UnU0-yLNCZ-OPr3/edit?usp=share_link&ouid=114060678686349206997&rtpof=true&sd=true'

spreadsheet_url2 = 'https://docs.google.com/spreadsheets/d/1j-xZhA0ozhsDRXf-0qoAqrGgHxF-dwpo/edit?usp=share_link&ouid=114060678686349206997&rtpof=true&sd=true'

spreadsheet1 = client.open_by_url(spreadsheet_url1)

spreadsheet2 = client.open_by_url(spreadsheet_url1)

# Read the contents of a specific sheet within the spreadsheet into a Pandas DataFrame
sheet_name = 'Sheet1'  # Replace with the desired sheet name
worksheet1 = spreadsheet1.worksheet(sheet_name)
worksheet2 = spreadsheet2.worksheet(sheet_name)
data1 = worksheet1.get_all_values()
data2 = worksheet2.get_all_values()

def readActivityDataFrame():
    # activities_df = pd.read_excel(spreadsheet_url1)
    return data1

def readUsersCompletionDataFrame():
    # completerates_df = pd.read_excel(spreadsheet_url2)
    return data2 

def updateUserCompletionRates(data):
    userId = int(data['userId'])
    activityId = int(data['activityId'])
    complete_score = int(data['complete_score'])
    satisfaction_score = float(data['satisfaction_score'])

    wb_append = load_workbook("Python-Recommendation-Model/src/data/completeratecheckappend.xlsx")

    sheet = wb_append.active
    row = (userId,activityId, complete_score,satisfaction_score)
    sheet.append(row)
    wb_append.save('Python-Recommendation-Model/src/data/completeratecheckappend.xlsx')
