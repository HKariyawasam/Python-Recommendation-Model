import pandas as pd
from flask import jsonify
from openpyxl import load_workbook

def readActivityDataFrame():
    activities_df = pd.read_excel('Python-Recommendation-Model/src/data/activities.xlsx')
    return activities_df

def readUsersCompletionDataFrame():
    completerates_df = pd.read_excel('Python-Recommendation-Model/src/data/completerate.xlsx')
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
