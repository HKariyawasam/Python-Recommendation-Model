import pandas as pd
from flask import jsonify
from openpyxl import load_workbook

def readActivityDataFrame():
    activities_df = pd.read_excel('C:/Users/Hasani Gayani/Desktop/MODEL TRAINING/Physical_Ac_Rec/activities.xlsx')
    return activities_df

def readUsersCompletionDataFrame():
    completerates_df = pd.read_excel('C:/Users/Hasani Gayani/Desktop/MODEL TRAINING/Physical_Ac_Rec/completerate.xlsx')
    return completerates_df

def updateUserCompletionRates(data):
    userId = int(data['userId'])
    activityId = int(data['activityId'])
    complete_score = int(data['complete_score'])
    satisfaction_score = float(data['satisfaction_score'])

    wb_append = load_workbook("C:/Users/Hasani Gayani/Desktop/MODEL TRAINING/Physical_Ac_Rec/completeratecheckappend.xlsx")

    sheet = wb_append.active
    row = (userId,activityId, complete_score,satisfaction_score)
    sheet.append(row)
    wb_append.save('C:/Users/Hasani Gayani/Desktop/MODEL TRAINING/Physical_Ac_Rec/completeratecheckappend.xlsx')
