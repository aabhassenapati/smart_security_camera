
from __future__ import print_function  
from xlrd import open_workbook
from googleapiclient.discovery import build  
from httplib2 import Http  
from oauth2client import file, client, tools  
from oauth2client.service_account import ServiceAccountCredentials 
import xlrd
import time
import datetime
import cv2
import xlwt


MY_SPREADSHEET_ID = '12EC6AmdEUOcwQgLIDZ64rMbR9pDLmvMdb6gjcN0w5mY'


def update_sheet(sheetname, name , time):  
    """update_sheet method:
       appends a row of a sheet in the spreadsheet with the 
       the latest temperature, pressure and humidity sensor data
    """
    # authentication, authorization step
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    creds = ServiceAccountCredentials.from_json_keyfile_name( 
            'client_secret.json', SCOPES)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Call the Sheets API, append the next row of sensor data
    # values is the array of rows we are updating, its a single row
    values = [ [ time, 
        'Person', name] ]
    body = { 'values': values }
    # call the append API to perform the operation
    result = service.spreadsheets().values().append(
                spreadsheetId=MY_SPREADSHEET_ID, 
                range='Sheet1' + '!A1:C1',
                valueInputOption='USER_ENTERED', 
                insertDataOption='INSERT_ROWS',
                body=body).execute()  


row=2
count = 2

while True:
    # A3 to D7
    
 workbook = xlrd.open_workbook(r"book2.xls")
 sheet = workbook.sheet_by_index(0)
 m = str(sheet.cell(1,2))
 o =int( float (m[7:]))
 n=o+2
 print(n)
       
 if(count<n):
    data = str(sheet.cell(row,0))
    time = str(sheet.cell(row,1))
    print(count)
    update_sheet("Face_Recognition" , data[5:] , time[5:])
    count=count+1
    row = row + 1
 
     #print (count)
   
     
