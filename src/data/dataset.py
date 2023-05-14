import pandas as pd
from flask import jsonify
from openpyxl import load_workbook


spreadsheet_url1 = 'https://docs.google.com/spreadsheets/d/10AQCTZt-mpp0xQ5T4UnU0-yLNCZ-OPr3/edit?usp=share_link&ouid=114060678686349206997&rtpof=true&sd=true'

spreadsheet_url2 = 'https://docs.google.com/spreadsheets/d/1j-xZhA0ozhsDRXf-0qoAqrGgHxF-dwpo/edit?usp=share_link&ouid=114060678686349206997&rtpof=true&sd=true'



def readActivityDataFrame():
    activities_df = pd.read_excel(spreadsheet_url1)
    return activities_df

def readUsersCompletionDataFrame():
    completerates_df = pd.read_excel(spreadsheet_url2)
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
