import random
import re
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from gspread.utils import rowcol_to_a1

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

# add credentials to the account
# authorize the clientsheet 
creds = ServiceAccountCredentials.from_json_keyfile_name("excelscraper-11d9cd7fa778.json", scope)
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
        
        if worksheet.get_tab_color() == None: # if the tab color is not set to black
            title = worksheet.title
            match = re.compile(r'\d+').search(title)# Search for the pattern in the text
            tableNumber = int(match.group())

            if tableNumber > start_sheet_index:
                filteredWorksheets.append(worksheet)
    
    return filteredWorksheets

# access to google sheets docs
milSTDFile = client.open('Copy of MIL-STD-1472H (10)')
allWorkMilSTDsheets = milSTDFile.worksheets() # get all worksheets

cleanDataFile = client.open('Copy of MilSTD1472HS5CleanData')
cleanDataSheet = cleanDataFile.get_worksheet(0) # get Clean Data worksheet


# nonProcessedSheets = filterWorksheets(allWorkMilSTDsheets) #use this to loop through all sheets in the whole document
# for currMilSTDSheet in nonProcessedSheets:


currMilSTDSheet = milSTDFile.get_worksheet(51) #TODO only processing 1 specific sheet in the document for testing, remove for full iteration later
readReqCount = 0;
writeReqCount = 0;

for i in range(1,currMilSTDSheet.row_count):
    #TODO add catch rate limit error here to sleep for 10 seconds and try again

    currCell = currMilSTDSheet.cell(col=1,row=i)
    try:
        currRowText = currCell.value # get the value at the specific cell
    except Exception as e:
        print(f"A read error limitation occurred: {str(e)}")

    readReqCount += 1 # update iterations of read reqs for usage limitations wait time calculations

    if currRowText == None:
        emptyCount += 1
    else: # row has text
        emptyCount = 0 # reset the emptyCount

        # Regular expression to match each entry
        pattern = re.compile(r'^\s*((?:\d+\.)+\d+)\s+([^\d\s].+?\.)\s*(.*?)(?=^\s*(?:\d+\.)+\d+\s+[^\d\s]|$)', re.DOTALL | re.MULTILINE)

        # Find all matches
        matches = pattern.findall(currRowText)

        # Extract and print the parts for each match
        for match in matches:
            section, term, description = match
            print(f"Section: '{section}'")
            print(f"Term: '{term}'")
            print(f"Description: '{description.strip()}'\n")
            print("=====================================")

            nextAvailRow = len(cleanDataSheet.get_all_values()) + 1

            try:
                cleanDataSheet.update_cell(nextAvailRow, 1, section)
            except Exception as e:
                print(f"A section write error limitation occurred: {str(e)}")

            try:
                cleanDataSheet.update_cell(nextAvailRow, 2, term)
            except Exception as e:
                print(f"A term write error limitation occurred: {str(e)}")

            try:
                cleanDataSheet.update_cell(nextAvailRow, 3, description.strip())
            except Exception as e:
                print(f"A description write error limitation occurred: {str(e)}")

            #TODO add catch of limit usage write requests and retrying after sleeping using "min(((2^n)+random_number_milliseconds), maximum_backoff)" where n = read or write req count, maximum_backoff = 64.
            writeReqCount += 3 #TODO possibly remove the equation for sleep and just set back off time to hardcode 10s before retrying
            
            currCella1Notation = rowcol_to_a1(i,1);
            format = CellFormat(
                backgroundColor=Color(1,1,0) #RGB values for yellow highlight
            )
            format_cell_range(currMilSTDSheet,currCella1Notation,format)
        # print("next row")

    if emptyCount == 5: # if the row is empty for 5 times, then go to next Table/Sheet
        break
currMilSTDSheet.update_tab_color('#000000') # change the tab color to black after the worksheet has been cleaned up