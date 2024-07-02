import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from gspread.utils import rowcol_to_a1
from definitionsScraper import exponential_backoff

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

# add credentials to the account
# authorize the clientsheet 
creds = ServiceAccountCredentials.from_json_keyfile_name("excelscraper-11d9cd7fa778.json", scope)
client = gspread.authorize(creds)
  
yellow_highlight = CellFormat( #definition for yellow highlight
    backgroundColor=Color(1,1,0) #RGB values for yellow highlight
)

# access to google sheets docs

cleanDataFile = client.open('Copy of MilSTD1472HS5CleanData')
cleanDataSheet = cleanDataFile.get_worksheet(0) # get Clean Data worksheet

request_attempt = 0

for i in range(2,cleanDataSheet.row_count):
    request_attempt = 0
    while True:
        try:
            currSectionText = cleanDataSheet.cell(row=i,col=1).value
            break
        except Exception as e:
            print(f"A read error limitation occurred: {str(e)}")
            request_attempt += 1
            exponential_backoff(request_attempt)
    
    request_attempt = 0
    while True:
        try:
            currCell_format = get_effective_format(cleanDataSheet, rowcol_to_a1(i,1)) # row = i, col = 1
            break
        except Exception as e:
            print(f"A read error limitation occurred: {str(e)}")
            request_attempt += 1
            exponential_backoff(request_attempt)
        
    if currCell_format: # if format exists, convert to proper object type for highlight comparison
        currCell_format_highlight = CellFormat(
        backgroundColor=Color(currCell_format.backgroundColor.red, currCell_format.backgroundColor.green, currCell_format.backgroundColor.blue)
    )


    if currSectionText == None:
        emptyCount += 1
    elif currCell_format_highlight == yellow_highlight: #has text with highlight checking to skip over processed cells
        continue
    else: # row has text and has not been processed
        emptyCount = 0 # reset the emptyCount

        # evaluate if subsection
        if not re.match(re.compile(r'^\d+(\.\d+)+$'), currSectionText):

            # extract the cell text + text in the description section

            while True:
                try:
                    currDescriptionText = cleanDataSheet.cell(row=i,col=3).value
                    break
                except Exception as e:
                    print(f"A read error limitation occurred: {str(e)}")
                    request_attempt += 1
                    exponential_backoff(request_attempt)

            while True:
                try:
                    prevDescriptionText = cleanDataSheet.cell(row = i - 1, col = 3).value
                    break
                except Exception as e:
                    print(f"A read error limitation occurred: {str(e)}")
                    request_attempt += 1
                    exponential_backoff(request_attempt)
            
            #append it to the row above's description
            appendSubsectionText = prevDescriptionText + '\n' + currSectionText +  " " + currDescriptionText
            request_attempt = 0
            while True:
                try:
                    cleanDataSheet.update_cell(row = i - 1, col = 3,value=appendSubsectionText)
                    break
                except Exception as e:
                    print(f"A write error limitation occurred: {str(e)}")
                    request_attempt += 1
                    exponential_backoff(request_attempt)
            
            print(appendSubsectionText)
            
            #highlight and delete the row
            request_attempt = 0
            while True:
                try:
                    cleanDataSheet.delete_rows(i)
                    break
                except Exception as e:
                    print(f"A write error limitation occurred: {str(e)}")
                    request_attempt += 1
                    exponential_backoff(request_attempt)
            

        else: # processed
            request_attempt = 0
            for j in range(1,4):
                while True:
                    try:
                        format_cell_range(cleanDataSheet, rowcol_to_a1(i,j), yellow_highlight)
                        break
                    except Exception as e:
                        print(f"A write error limitation occurred: {str(e)}")
                        request_attempt += 1
                        exponential_backoff(request_attempt)



    if emptyCount == 5: # if the row is empty for 5 times, then go to next Table/Sheet
        break

    #TODO fix iteration after processing a subsection
    #if not, then after deletion, check the same index again
    #use a while loop instead with a index counter for manual iteration manipulation