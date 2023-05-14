import pandas as pd
from flask import jsonify
from openpyxl import load_workbook
import os


import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Define the scope and credentials for accessing the Google Drive and Sheets APIs
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_json = {
  "type": "service_account",
  "project_id": "rec-engine-packages",
  "private_key_id": "7402e6fc1321a49c5a8a5f0dd7fd0ecebd50ac0a",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC9Rj2lbyFHiwZY\nwHSquOsBwUUqUafrpY76Dn5b6WfWS+jB6O4T1HDGOFlpyWj0FLZv8ZEW2JOiODYm\nrD3baj6etJzwCN63b1ZgwPFbYZrsu93Ooo5soad7uMlv78QTfHlbF5xqtsi1ZX9z\n9QI0F0sXpry1y+iiTlzAgGn4LxbmPhSUbmj4EKjOd8SSgXAuQ/PcjuRhkcUii11U\nM88E8C+lbAnufsYeZJkvw/oUiO0hHjgDGUc4baJG5rOOvusP0Syh65W8+la8GKfe\nXYFlTJPGKvXR4APfju48F6RzCqFBY3h+cyd/Auub+aXU3Ipg4j578iaQ0SSaWssV\nBRgQmsDLAgMBAAECggEAXdoE+lGi9lcs4/g2QWsk+H9NnQDVW2cCAOML6+ZN+zgz\np2cpGkMWUmuhhm98776Pge2N7H6ioebujvmMSI6jk01qnik/ykRWV+6EHzREPkuf\nXfVD2wDDS/liRPNmTeVERZNtF4sf9bZo3uYn3A2KyiT+4MRFy6lAm6FxKkOrPqnB\nirIlEfm6sdJd6a1sNwP/VPjlTChllYiTUGBu0ohhQPKyOQcXGdTeIzruLSs1IIwP\nRYa1efn8+N+vdj+dBLuNtKLrBimYpxtJslfeDtxOoLVR+dyKf/2ySuX/+hdu8Xpd\nbJS4YQTR65o30ER15yyNqonlceYwwOAfFPU6+uskWQKBgQDxmO6F8JnYZGICOVLH\nLYcnSyjyAzBwUZgMmH+wN7GDHddRrbJr5NbIVjamjpK/7OM4mhsslAbO/xpV8jzj\nQ0okD2FKOs7sBuMFvz4rb3dnsoXoCDq4tNYcsl/rAcwy+Yl0Az0VJ6QsMSBdBOk1\nfC3GtyUcazz4SKEMJhp/eoxXwwKBgQDIjsvx0poW5XtfaT0Lu0YI9pASLb7IjZQY\nEJmzXW4ZjSutGrPsJjrkBQRiPfu8Z3ZKii2IPH19dQkAjkKo2XBGbFoL2PpN57Xe\ng/YYr2gPqlnR1yDbiOAiXPc3XSb+f+5Zzk0GGvAq/SxgOaP2M5XeLtjZOjmyGlyQ\nXv6zWojqWQKBgBHiAw3RAGI/E+4cNh9eJFnpO0+mosg9keakxxbRGIefBtgZ5lIM\nL0XG8+aiOQSR0UPYTFihYFukEFv6QT2FNpCyvr3S2own+lfrjvuCFbGMSlMhgM85\nO3wGTAlGKcpAJEd9EeFl/MX0oPOhsG1wEqdZ2RpgabPrFFik3WNAO/EnAoGAecHP\naNkteRcKhksSp4ujIg/mOVMSTxk8vjtdxHnFPLfquyXJ82TmPcYZ+jadHK1HMEuh\nBuOKX97sfyzepTmUovxm0miA1UkFrbg5cJUUvOXzr6RGK2F2iQYdg7wGz10Fa/oF\n4t35zm9zQFveAbshkgio14A0xL6iUXeKc4JUOskCgYEA1BXsxXbQ3y8CMLi1ohpg\n1yFtvURMLvgg8EnZq84UvkIFL5rO51lzoY3s52oaTB/JxtMViwnf3iZ/5dkUCQAZ\nMo1YOynVe0Go1u2Oq4iY7EyToG4K9aXn8NA/kOQV8YjZewtUAQvriS6G2kyCUJt8\nRcGWuVQ3lKZTWrVNIBmGHn8=\n-----END PRIVATE KEY-----\n",
  "client_email": "pa-engine@rec-engine-packages.iam.gserviceaccount.com",
  "client_id": "106200314142575525011",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/pa-engine%40rec-engine-packages.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_json, scope)


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
