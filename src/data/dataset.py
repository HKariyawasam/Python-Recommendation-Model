import pandas as pd
from flask import jsonify
from openpyxl import load_workbook
import os
import requests


# Define the shared link URL of the Google Sheets document
shared_link_url1 = "https://docs.google.com/spreadsheets/d/10AQCTZt-mpp0xQ5T4UnU0-yLNCZ-OPr3/export?format=csv"
shared_link_url2 = "https://docs.google.com/spreadsheets/d/1j-xZhA0ozhsDRXf-0qoAqrGgHxF-dwpo/export?format=csv"
# Send a GET request to download the CSV data
response1 = requests.get(shared_link_url1)
response2 = requests.get(shared_link_url2)







def readActivityDataFrame():
    # activities_df = pd.read_excel(spreadsheet_url1)
    if response1.status_code == 200:
        buffer1 = response1.content
        activities_df = pd.read_excel(buffer1, engine="openpyxl")
        return activities_df

def readUsersCompletionDataFrame():
    # completerates_df = pd.read_excel(spreadsheet_url2)
    if response2.status_code == 200:
        buffer2 = response2.content
        completerates_df = pd.read_excel(buffer2, engine="openpyxl")
        return completerates_df 

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
