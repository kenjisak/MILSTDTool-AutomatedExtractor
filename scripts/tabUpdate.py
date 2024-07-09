import random
import re
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

# add credentials to the account
# authorize the clientsheet 
creds = ServiceAccountCredentials.from_json_keyfile_name("../resources/excelscraper-11d9cd7fa778.json", scope)
client = gspread.authorize(creds)

# Function to perform exponential backoff
def exponential_backoff(attempt):
    maximum_backoff_time = 64;
    backoff_time = min(((2 ** attempt) + random.random()), maximum_backoff_time)  # Calculate backoff time with a maximum time in seconds set
    print(f"Retrying in {backoff_time} seconds...")
    time.sleep(backoff_time)
    
def filterWorksheets(allWorksheets):
    filteredWorksheets = []
    start_sheet_index = 50
    for worksheet in allWorksheets: # filter out the worksheets that have not been cleaned up *start from Table 51, unsure why some under Table 51 are not changed to black
        #print(worksheet.title, worksheet.get_tab_color())
        while True:
            try:
                currSheetTabColor = worksheet.get_tab_color()
                break
            except Exception as e:
                print(f"A read error limitation occurred: {str(e)}")
                request_attempt += 1
                exponential_backoff(request_attempt)

        if currSheetTabColor == None: # if the tab color is not set to black
            title = worksheet.title
            match = re.compile(r'\d+').search(title)# Search for the pattern in the text
            tableNumber = int(match.group())

        if tableNumber > start_sheet_index:
            filteredWorksheets.append(worksheet)
    
    return filteredWorksheets

# access to google sheets docs
milSTDFile = client.open('Copy of MIL-STD-1472H (10)')
allWorkMilSTDsheets = milSTDFile.worksheets() # get all worksheets

nonProcessedSheets = filterWorksheets(allWorkMilSTDsheets) #use this to loop through all sheets in the whole document
for currMilSTDSheet in nonProcessedSheets:
    request_attempt = 0

    while True:
        try:
            currMilSTDSheet.update_tab_color('#000000') # change the tab color to black after the worksheet has been cleaned up
            break
        except Exception as e:
            print(f"A write error limitation occurred: {str(e)}")
            request_attempt += 1
            exponential_backoff(request_attempt)