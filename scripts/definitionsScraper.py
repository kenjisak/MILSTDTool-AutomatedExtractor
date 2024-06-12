import gspread
from oauth2client.service_account import ServiceAccountCredentials

import re

# define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']

# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name("excelscraper-11d9cd7fa778.json", scope)
# authorize the clientsheet 
client = gspread.authorize(creds)

# get the instance of the Spreadsheet
sheet = client.open('Copy of MIL-STD-1472H (10)')

worksheet = sheet.get_worksheet(50)
# worksheet.update_cell(1, 1, "I just wrote to a spreadsheet using Python!")

# if rows are empty for 5 times, then stop the loop
currRowText = worksheet.cell(col=1,row=10).value # get the value at the specific cell


# Regular expression to match each entry
pattern = re.compile(r'((?:\d+\.)+\d+)\s+([^\d\s].+?\.)\s*(.*?)(?=(?:\d+\.\d+)|$)', re.DOTALL)

# Find all matches
matches = pattern.findall(currRowText)

# Extract and print the parts for each match
for match in matches:
    part1, part2, part3 = match
    print(f"Part 1: '{part1}'")
    print(f"Part 2: '{part2}'")
    print(f"Part 3: '{part3.strip()}'\n")