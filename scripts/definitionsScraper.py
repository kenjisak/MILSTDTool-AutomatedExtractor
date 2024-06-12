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

# get the instance of the Spreadsheet
sheet = client.open('Copy of MIL-STD-1472H (10)')
allWorksheets = sheet.worksheets() # get all worksheets

filterWorksheets(allWorksheets)


worksheet = sheet.get_worksheet(70)
# worksheet.update_cell(1, 1, "I just wrote to a spreadsheet using Python!")

for i in range(1,worksheet.row_count):
    #add catch rate limit error here to sleep for 10 seconds and try again
    currRowText = worksheet.cell(col=1,row=i).value # get the value at the specific cell
    
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
            part1, part2, part3 = match
            print(f"Part 1: '{part1}'")
            print(f"Part 2: '{part2}'")
            print(f"Part 3: '{part3.strip()}'\n")


    if emptyCount == 5: # if the row is empty for 5 times, then go to next Table
        break