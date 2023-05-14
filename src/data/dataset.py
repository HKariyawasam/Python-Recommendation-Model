import pandas as pd
from flask import jsonify
from openpyxl import load_workbook

import gdown

# Shared Google Drive link of the Excel file
shared_link = 'https://docs.google.com/spreadsheets/d/1xjEEOrjQkUBLDtb1UHGrurKkVdmp0meS/edit?usp=share_link&ouid=114060678686349206997&rtpof=true&sd=true'
shared_link2 = 'https://docs.google.com/spreadsheets/d/18lfP1BZOfpy1S1Rjgrr1G7sy0JiknDQV/edit?usp=sharing&ouid=114060678686349206997&rtpof=true&sd=true'

# Extract the file ID from the shared link
file_id = shared_link.split('/')[-2]
file_id2 = shared_link2.split('/')[-2]

# Construct the download link for the file
download_link = f'https://drive.google.com/uc?id={file_id}'
download_link2 = f'https://drive.google.com/uc?id={file_id2}'

# Download the Excel file
file_path = 'file.xlsx'
file_path2 = 'file2.xlsx'
gdown.download(download_link, file_path, quiet=False)
gdown.download(download_link, file_path2, quiet=False)








def readActivityDataFrame():
    activities_df = pd.read_excel(file_path2)
    return activities_df

def readUsersCompletionDataFrame():
    completerates_df = pd.read_excel(file_path)
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
