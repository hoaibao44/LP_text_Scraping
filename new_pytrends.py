from __future__ import print_function
import json
import sys
import time
from pprint import pprint
from datetime import date, datetime, timedelta
from pytrends.request import TrendReq
from time import sleep

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import openpyxl
from openpyxl import Workbook, load_workbook

import pickle
import os.path
import requests
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


min_value = 5000
now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

def get_Gsheet_info(sheet_id,sheet_range):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # The ID and range of a sample spreadsheet.
    SPREADSHEET_ID = sheet_id
    RANGE_NAME = sheet_range
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    return values

def KW_trend_info(in_info,inKW):
    rows_data =[]
    rising_df = in_info
    if rising_df is not None:
        i=0
        for i in range(0,rising_df['value'].count()-1):
            if rising_df.loc[i]['value'] >= min_value and rising_df.loc[i]['value'] < 5000:
                #print(rising_df.loc[i]['query'])
                #CW_msg_query =  CW_msg_query +'└\t'+ rising_df.loc[i]['query'] +'（'+ str(rising_df.loc[i]['value']) +'%増加）\n'
                rows_data.append([now,inKW,rising_df.loc[i]['query'],str(rising_df.loc[i]['value']) +'%増加'])
            elif rising_df.loc[i]['value'] >= 5000:
                #CW_msg_query =  CW_msg_query +'└\t'+ rising_df.loc[i]['query'] +'（急激増加）\n'
                rows_data.append([now,inKW,rising_df.loc[i]['query'],'急激増加'])
        if len(rows_data)==0:
            return ['no_data',[]]
        else:
            return ['has_data',rows_data]
    else:
        return ['no_result',[]]
        
def addRow_to_Gsheet(sheet_id,sheet_range,in_array):
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # The ID and range of a sample spreadsheet.
    SPREADSHEET_ID = sheet_id
    RANGE_NAME = sheet_range    

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    #print(values[0])

    values = in_array
    body = {
    "majorDimension": "ROWS",
    "values": values
    }
    result = sheet.values().append(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME,
    valueInputOption="USER_ENTERED", body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

def get_last_row(sheet_id,sheet_range):

    mid_value = get_Gsheet_info(sheet_id,sheet_range)
    last_row = len(mid_value)
    return last_row

def update_query_data():
   
    KW_list =get_Gsheet_info('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Gtrend_KW!A1:A')
    print('All KWs: '+str(len(KW_list)))

    KW_list[0].append('result')
    i=1
    for mykw in KW_list[1:] :
        try:
            pytrends = TrendReq(hl='ja-JP', tz=360,timeout=(10,25),geo = 'JP',proxies=['https://172.19.119.54:3128'], retries=3)
            pytrends.build_payload(mykw, timeframe='now 1-H', geo='JP')
        except:
            pytrends = TrendReq(hl='ja-JP', tz=360,timeout=(10,25),geo = 'JP',proxies=['https://10.141.64.176:3128'], retries=3)
            pytrends.build_payload(mykw, timeframe='now 1-H', geo='JP')           
        df = pytrends.related_queries()
        rising_df = df[mykw[0]]['rising']

        print(rising_df)
        
        myresult = KW_trend_info(rising_df,mykw[0])

        KW_list[i].append(myresult[0])
        KW_list[i].append(myresult[1])
        if myresult[0] == 'has_data':
            print('KW no.'+str(i)+': '+mykw[0] +' -- '+str(len(myresult[1])))
            addRow_to_Gsheet('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Log!A1:D',myresult[1])
        else:
            print('KW no.'+str(i)+': '+mykw[0] +' -- '+myresult[0])
        i=i+1
    
    final_resutl = []
    for myKW in KW_list[1:]:
        if myKW[1]=='has_data':
            final_resutl.append(myKW)
    print(final_resutl)

def sendMessage(room_id, text_message_as_string):
    APIKEY = '66d7468c0232a7a8d0c24d8ef8c2a71c'
    ENDPOINT = 'https://api.chatwork.com/v2'
    room_id = room_id
    bodytext = text_message_as_string
    post_message_url = '{}/rooms/{}/messages'.format(ENDPOINT, room_id)
         
    headers = { 'X-ChatWorkToken': APIKEY }
    #print(str(bodytext))
    params = { 'body': str(bodytext) }
         
    resp = requests.post(post_message_url,
                         headers=headers,
                         params=params)
    print(resp.encoding)
    #print(resp.url)
    return pprint(resp.content)

def sendFile(room_id,file_name,file_path):
    APIKEY = '66d7468c0232a7a8d0c24d8ef8c2a71c'
    ENDPOINT = 'https://api.chatwork.com/v2'
    room_id = room_id
    file_path = file_path
    file_name = file_name
    post_message_url = '{}/rooms/{}/files'.format(ENDPOINT, room_id)
         
    headers = { 'X-ChatWorkToken': APIKEY }

    files  = {'file': (file_name,open(file_path, 'rb'),'application/vnd.ms-excel')}
         
    resp = requests.post(post_message_url,headers=headers,files=files)
    #print(resp.url)
    return pprint(resp.content)

if __name__ == '__main__':
    #get last row in sheet before runing
    last_row_start = get_last_row('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Log!A1:A')
    #last_row_start = 48
    print(str(last_row_start))

    update_query_data()

    #get last row in sheet after updated
    last_row_end = get_last_row('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Log!A1:A')
    print(str(last_row_end))
    
    query_info = get_Gsheet_info('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Log!A'+str(last_row_start+1)+':D'+str(last_row_end))

    anken_all = ['C','D','E','F']

    for myClient in anken_all:
        anken_info = get_Gsheet_info('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y','Gtrend_KW!'+myClient+'1:'+myClient)
        client_name = anken_info[0][0]
        
        #TO CW
        CW_ID = anken_info[1][0]
        print(client_name)
        print(CW_ID)
        
        CW_msg = ''

        for myKW in anken_info[3:]:
            has_data = False
            CW_sub_msg = ''
            #print('processing:'+myKW[0])
            for row in query_info:
                if row[1] == myKW[0] and row[3] == '急激増加':
                    CW_sub_msg = CW_sub_msg + '\n└	' + row[2]+'（'+ row[3]+'）'
                    has_data = True
            if has_data:
                CW_msg = CW_msg + '▼Rising query for KW:  '+myKW[0]+CW_sub_msg+'[hr]'
        
        if CW_msg != '' and client_name != 'U-Next':
            CW_msg = anken_info[2][0].format(Run_time=now)+CW_msg+'[/info]'
            sendMessage(CW_ID,CW_msg)
        elif CW_msg == '' and client_name != 'U-Next':
            pass
            sendMessage(CW_ID,'[info]データなし[/info]')
        print(CW_msg)

        #update to G sheet
        allRows = []
        for row in query_info:
            has_data = False
            sub_allRows = []
            for myKW in anken_info[3:]:
                if row[1] == myKW[0] and row[3] == '急激増加':
                    sub_allRows.append(row[0])
                    sub_allRows.append(row[1])
                    sub_allRows.append(row[2])
                    sub_allRows.append(row[3])
                    has_data = True
            if has_data:
                allRows.append(sub_allRows)
        if len(allRows)!=0:
            addRow_to_Gsheet('1coCs1viBARfhjqdLhwkvPzg3oGOtAVGzBGXWx07h94Y',client_name+'!A1:D',allRows)
    print('※DONE※')
