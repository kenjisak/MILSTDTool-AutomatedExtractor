import re
from sortedcontainers import SortedSet
import camelot

import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("excelscraper-11d9cd7fa778.json", scope)
client = gspread.authorize(creds)

cleanDataFile = client.open('Copy of MilSTD1472HS5CleanData')

milstdpdf_file_path = '../resources/MIL-STD-1472H.pdf'
tabledata_file_path = '../resources/contentsData/tableData.txt'
figuredata_file_path = '../resources/contentsData/figureData.txt'


# Function to extract integers from each line in a file
def extract_pageNumbers_from_file(txt_file_path):
    # pageNumberList = list()
    pageNumberSet = SortedSet()

    with open(txt_file_path, 'r') as file:
        for line in file:
            pageNumbers = re.findall(r'\d+', line)
            pageNumberSet.add(int(pageNumbers[-1]))
            # pageNumberList.append(int(pageNumbers[-1]))
    # return pageNumberSet, pageNumberList

    # print(len(pageNumberSet))
    # print(pageNumberList)
    return pageNumberSet

def extract_titles_from_file(file_path):
    extracted_lines = []
    with open(file_path, 'r') as file:
        for line in file:
            # Use regex to match text before the trailing dots and digits
            match = re.match(r'^(.*?)(\s*\.*\d*\s*)$', line)
            if match:
                extracted_lines.append(match.group(1).strip())

    # for text in extracted_lines:
    #     print(text)

    return extracted_lines

def upload_csv_file(csv_file_path):
    # Extract the name of the CSV file without the extension
    new_sheet_name = os.path.splitext(os.path.basename(csv_file_path))[0]

    # Read data from the CSV file
    with open(csv_file_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        data = list(reader)

    num_rows = len(data) #inserts a row so dont need to set the size to exact
    num_cols = max(len(row) for row in data)

    # Check if the sheet already exists
    try:
        worksheet = cleanDataFile.worksheet(new_sheet_name)
        print(f"Sheet '{new_sheet_name}' already exists.")
    except gspread.exceptions.WorksheetNotFound:
        # Create a new sheet with the name derived from the CSV file name
        worksheet = cleanDataFile.add_worksheet(title=new_sheet_name, rows=num_rows, cols=num_cols)
        print(f"New sheet '{new_sheet_name}' created.")

    # Clear the existing content in the worksheet
    worksheet.clear()

    # Write the data to the new Google Sheet
    for row_index, row in enumerate(data, start=1):
        worksheet.insert_row(row, row_index)

    worksheet.resize(rows=num_rows, cols=num_cols) # resize sheet to be exact
    print(f"Data written to sheet '{new_sheet_name}' successfully.")

# TODO add figure/image extraction plus table extraction for that same page to be paired with
# TODO add api update requests to google sheets doc after each table
# TODO use a sheet storage for persistence in pages alreayd processed for fast processing
# TODO ***** add checking of first row for continuation of tables
# TODO add function for retrying requests to reduce duplication

# upload_csv_file("../resources/tablesCSV/" + "TABLE I. Mechanical control criteria.csv")

tables_page_numbers = extract_pageNumbers_from_file(tabledata_file_path)
figures_page_numbers = extract_pageNumbers_from_file(figuredata_file_path)

tables_page_titles = extract_titles_from_file(tabledata_file_path)
figures_page_titles = extract_titles_from_file(figuredata_file_path)

page_offset = 17 # page 1 starts after page 17, page offset + extracted page = correct page number
title_index_counter = 0

for i in range(len(tables_page_numbers)):
    corrected_page_number = tables_page_numbers[i] + page_offset
    currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(corrected_page_number))
    # while True:
    #     currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(47 + page_offset))

    for j in range(len(currTables)):
        # print(currTable[i].df)
        print("Page Number: " + str(tables_page_numbers[i]) + str(currTables[j].parsing_report))
        currTables[j].to_csv("../resources/tablesCSV/" + tables_page_titles[title_index_counter] + '.csv')
        title_index_counter += 1
        # if currTables[j].parsing_report['accuracy'] >= 97:
        #     break