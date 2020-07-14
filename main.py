from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import selenium.webdriver.support.expected_conditions as EC
import datetime
from time import sleep
from openpyxl import load_workbook
import os
from glob import glob
import requests

import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

def highlight(element):
    """Highlights (blinks) a Selenium Webdriver element"""

    browser = element._parent

    def appy_style(s):
        browser.execute_script("arguments[0].setAttribute('style', arguments[1]);", element, s)

    original_style = element.get_attribute('style')
    appy_style("background: yellow; border: 2px solid red;")
    sleep(.2)
    appy_style(original_style)

class Driver:

    def __init__(self, driver):
        self.driver = driver

    def get(self, url):
        self.driver.get(url)

    def click(self, xpath, *arg):
        if len(arg) == 0:
            arg = 30
        else:
            arg = arg[0]
        WebDriverWait(self.driver, arg).until(EC.visibility_of_element_located((By.XPATH, xpath))).click()
        sleep(0.3)

    def send_keys(self, xpath, string, *args):
        if len(args) == 0:
            arg = 30
        else:
            arg = args[0]
        WebDriverWait(self.driver, arg).until(EC.visibility_of_element_located((By.XPATH, xpath))).send_keys(string)
        sleep(0.3)

    def clear(self, xpath, *args):
        if len(args) == 0:
            arg = 30
        else:
            arg = args[0]
        WebDriverWait(self.driver, arg).until(EC.visibility_of_element_located((By.XPATH, xpath))).clear()
        sleep(0.3)

    def highlight(self, xpath):
        """Highlights (blinks) a Selenium Webdriver element"""

        element = WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, xpath)))

        browser = element._parent

        def appy_style(s):
            browser.execute_script("arguments[0].setAttribute('style', arguments[1]);", element, s)

        original_style = element.get_attribute('style')
        appy_style("background: yellow; border: 2px solid red;")
        sleep(.3)
        appy_style(original_style)
        sleep(0.3)

    def wait_for_all_elements(self, xpath, timeout):
        elements = list()
        for second in range(0, timeout):

            elements = self.driver.find_elements_by_xpath(xpath)

            if len(elements) > 0:
                return elements

            sleep(1)

        return elements

    def quit(self):
        self.driver.quit()
        
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

def get_page_text(in_url_table):
    in_url_table = in_url_table
    global chrome
    global driver
    master_tag =['select','p']
    str_output_text ={'url_stt':[],'url_text':[]}
    out_data =[]
    chrome = webdriver.Chrome(executable_path=os.getcwd()+ '/chromedriver.exe')
    driver = Driver(chrome)

    #for each row in data file, take url and filter setting
    for i in range(0,len(in_url_table)):
        print('Processing: '+in_url_table[i][0])
        stt_array =[]
        stt_value ='OK'
        
        try:
            driver.get(in_url_table[i][0])
            sleep(5)
            #get all elements of page
            
            elems =WebDriverWait(chrome, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*')))
            print('waiting for load completed')
            
            #for element in all elements, check if has text and extract
            out_elems =[]
            error_elems =[]
            
            for e in elems:
                #check if contains text
                isHastext = False
                try:

                    if e.text!='':
                        node_text= e.text
                        #node has no child
                        if len(e.find_elements_by_xpath('./*')) == 0 and out_elems.count(e.find_element_by_xpath('./parent::*'))==0:
                            parrent_node_text = e.find_element_by_xpath('./parent::*').text
                            isHastext = True
                            #if node_text.strip() == parrent_node_text.strip() or master_tag.count(e.tag_name) != 0:
                            #    isHastext = True
                        #node has 1 child
                        elif len(e.find_elements_by_xpath('./*')) == 1 and out_elems.count(e.find_element_by_xpath('./parent::*'))==0:
                            parrent_node_text = e.find_element_by_xpath('./parent::*').text
                            child_node_text = e.find_element_by_xpath('./*').text
                            if node_text.strip() != child_node_text.strip():
                                isHastext = True
                        #node has more than 1 child
                        elif len(e.find_elements_by_xpath('./*')) > 1 and out_elems.count(e.find_element_by_xpath('./parent::*'))==0:
                            if master_tag.count(e.tag_name) != 0:
                                isHastext = True

                except:
                    error_elems.append(e)
                
                #extract text and tag value
                if isHastext:
                    out_elems.append(e)
                    sub_data =[]
                    highlight(e)
                    print(' -- Text -- {b}'.format(b=e.text))
                    #print(e.rect)
                    #print(' -- Parent -- '+e.find_element_by_xpath('./parent::*').tag_name +' > '+e.find_element_by_xpath('./parent::*').get_attribute('class'))
                    #print(' -- Current -- '+e.tag_name +' > '+e.get_attribute('class'))
                    sub_data.append(process_start.strftime('%Y%m%d'))
                    sub_data.append(in_url_table[i][0])
                    sub_data.append(e.text)
                    sub_data.append(e.find_element_by_xpath('./parent::*').tag_name +' > '+e.find_element_by_xpath('./parent::*').get_attribute('class'))
                    sub_data.append(e.tag_name +' > '+e.get_attribute('class'))
                    out_data.append(sub_data)               
            print('Processing OK: '+in_url_table[i][0])
            print('Has {a} undetected elems and {b} elems have text'.format(a=str(len(error_elems)),b=str(len(out_elems))))

        except:
            stt_value ='Faulted'
            print('Processing Faulted: '+in_url_table[i][0])

        stt_array.append(in_url_table[i][0])
        stt_array.append(stt_value)
        stt_array.append(process_start.strftime('%Y%m%d'))
        str_output_text['url_stt'].append(stt_array)

    str_output_text['url_text'].append(out_data)

    chrome.close()
    return str_output_text

if __name__ == "__main__":
    import time
    global process_start
    process_start = datetime.datetime.now()

    sheet_id ='1JNeAwqA8snFqvZl4DXgrU3SXsVHrV62j-9cXgIFPoSw'
    in_url_table = get_Gsheet_info(sheet_id,'input_Url!A2:A') 
    
    output_text_data = get_page_text(in_url_table)

    output_url_stt_lastrow = get_last_row(sheet_id,'Output_STT!A1:A')
    addRow_to_Gsheet(sheet_id,'Output_STT!A{a}:C{b}'.format(a=str(output_url_stt_lastrow+1),b=str(output_url_stt_lastrow+len(output_text_data['url_stt']))),output_text_data['url_stt'])
    
    output_url_stt_lastrow = get_last_row(sheet_id,'Output_Text!A1:A')
    addRow_to_Gsheet(sheet_id,'Output_Text!A{a}:E{b}'.format(a=str(output_url_stt_lastrow+1),b=str(output_url_stt_lastrow+len(output_text_data['url_text'][0]))),output_text_data['url_text'][0])
    
    print('***DONE***')
