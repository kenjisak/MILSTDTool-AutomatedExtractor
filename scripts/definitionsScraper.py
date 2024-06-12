import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

# add credentials to the account
# authorize the clientsheet 
creds = ServiceAccountCredentials.from_json_keyfile_name("excelscraper-11d9cd7fa778.json", scope)
client = gspread.authorize(creds)


def filterWorksheets(allWorksheets):
    filteredWorksheets = []

    for worksheet in allWorksheets: # filter out the worksheets that have not been cleaned up *start from Table 51, unsure why some under Table 51 are not changed to black
        #print(worksheet.title, worksheet.get_tab_color())
        
        if worksheet.get_tab_color() == None: # if the tab color is not set to black
            title = worksheet.title
            match = re.compile(r'\d+').search(title)# Search for the pattern in the text
            tableNumber = int(match.group())

            if tableNumber > 50:
                filteredWorksheets.append(worksheet)
    
    return filteredWorksheets

# access to google sheets docs
milSTDFile = client.open('Copy of MIL-STD-1472H (10)')
allWorkMilSTDsheets = milSTDFile.worksheets() # get all worksheets

cleanDataFile = client.open('Copy of MilSTD1472HS5CleanData')
cleanDataSheet = cleanDataFile.get_worksheet(0) # get Clean Data worksheet


# filterWorksheets(allWorkMilSTDsheets)

milSTDSheet = milSTDFile.get_worksheet(51)
for i in range(1,milSTDSheet.row_count):
    #TODO add catch rate limit error here to sleep for 10 seconds and try again
    #TODO add highlight for the row that has been processed in MILSTD sheet and placed in CleanData sheet

    currRowText = milSTDSheet.cell(col=1,row=i).value # get the value at the specific cell
    
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

            cleanDataSheet.update_cell(nextAvailRow, 1, section)
            cleanDataSheet.update_cell(nextAvailRow, 2, term)
            cleanDataSheet.update_cell(nextAvailRow, 3, description.strip())
        print("next row")
    

    if emptyCount == 5: # if the row is empty for 5 times, then go to next Table
        break
milSTDSheet.update_tab_color('#000000') # change the tab color to black after the worksheet has been cleaned up