import re
from sortedcontainers import SortedSet
import camelot
import pymupdf
from camelot.parsers import Lattice

import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from definitionsScraper import exponential_backoff

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("../resources/excelscraper-11d9cd7fa778.json", scope)
client = gspread.authorize(creds)

cleanDataFile = client.open('Copy of MilSTD1472HS5CleanData')

milstdpdf_file_path = '../resources/MIL-STD-1472H.pdf'
tabledata_file_path = '../resources/contentsData/tableData.txt'
figuredata_file_path = '../resources/contentsData/figureData.txt'
saved_tables_csv_filepath = "../resources/tablesCSV"

page_offset = 17 # page 1 starts after page 17, page offset + extracted page = correct page number

def usage_limit_retry(func):
    request_attempt = 0

    while True:
        try:
            return func()
        except gspread.exceptions.APIError as e:
            print(f"A usage error limitation occurred: {str(e)}")
            request_attempt += 1
            exponential_backoff(request_attempt)
        except gspread.exceptions.WorksheetNotFound:
            # Reraise this exception so it can be caught by the calling function
            raise

# Function to extract integers from each line in a file
# def extract_pageNumbers_from_file(txt_file_path):
#     # pageNumberList = list()
#     pageNumberSet = SortedSet()

#     with open(txt_file_path, 'r') as file:
#         for line in file:
#             pageNumbers = re.findall(r'\d+', line)
#             pageNumberSet.add(int(pageNumbers[-1]))
#             # pageNumberList.append(int(pageNumbers[-1]))
#     # return pageNumberSet, pageNumberList

#     # print(len(pageNumberSet))
#     # print(pageNumberList)
#     return pageNumberSet

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
        worksheet = usage_limit_retry(lambda: cleanDataFile.worksheet(new_sheet_name))
        print(f"Sheet '{new_sheet_name}' already exists.")
    except gspread.exceptions.WorksheetNotFound:
        # Create a new sheet with the name derived from the CSV file name
        worksheet = usage_limit_retry(lambda: cleanDataFile.add_worksheet(title=new_sheet_name, rows=num_rows, cols=num_cols))
        print(f"New sheet '{new_sheet_name}' created.")

    # Clear the existing content in the worksheet
    usage_limit_retry(lambda: worksheet.clear())

    # Write the data to the new Google Sheet
    for row_index, row in enumerate(data, start=1):
        usage_limit_retry(lambda: worksheet.insert_row(row, row_index))

    usage_limit_retry(lambda: worksheet.resize(rows=num_rows, cols=num_cols)) # resize sheet to be exact
    print(f"Data written to sheet '{new_sheet_name}' successfully.")

def extract_tables():
    tables_page_numbers = extract_page_numbers(28,431)  # fixed missing page numbers for tables
    # figures_page_numbers = extract_pageNumbers_from_file(figuredata_file_path)

    tables_page_titles = extract_titles_from_file(tabledata_file_path)
    # figures_page_titles = extract_titles_from_file(figuredata_file_path)

    title_index_counter = 0

    previous_table_row_definitions = []
    
    for i in range(len(tables_page_numbers)): 
        corrected_page_number = tables_page_numbers[i] + page_offset
        currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(corrected_page_number))

        for j in range(len(currTables)):
            # print(currTable[j].df)
            print("Page Number: " + str(tables_page_numbers[i]) + str(currTables[j].parsing_report))

            currTable_title = corresponding_table_title_extraction(currTables[j])
            currTable_filepath = save_csv_incremental(currTables[j], saved_tables_csv_filepath,currTable_title)
            # upload_csv_file(currTable_filepath)

            # TODO refactor to be cleaner code and put into its own function for evaluation, find out other way by using nearest text/title for continued table evaluation * since doesnt always save properly for same first row
            # if previous_table_row_definitions == currTables[j].df.iloc[0].tolist(): #continued table
            #     currTable_filepath = save_csv_incremental(currTables[j], saved_tables_csv_filepath,tables_page_titles[title_index_counter - 1])
            #     upload_csv_file(currTable_filepath)

            # else:
            #     currTable_filepath = save_csv_incremental(currTables[j], saved_tables_csv_filepath, tables_page_titles[title_index_counter])
            #     upload_csv_file(currTable_filepath)
            #     title_index_counter += 1

            # previous_table_row_definitions = currTables[j].df.iloc[0].tolist()

def save_csv_incremental(table, directory, base_filename):
    # Create an incremental filename
    counter = 1
    path_and_filename = os.path.join(directory, f"{base_filename}.csv")
    
    # Check if the filename already exists and create a unique filename if necessary
    while os.path.exists(path_and_filename):
        path_and_filename = os.path.join(directory, f"{base_filename}_{counter}.csv")
        counter += 1
    
    # Save the table as a CSV file
    table.to_csv(path_and_filename)
    print(f"Table saved as {path_and_filename}")

    return path_and_filename

def extract_page_numbers(start, end): # fixed missing page numbers for tables in between by using range of start and end tables number
    return list(range(start, end + 1))

def corresponding_table_title_extraction(table):
    # TODO use references of order in camelot of page and also text that starts with TABLE
    # TODO add recognition of figure tables to ignore them and not save as csv, by checking matches order of FIGURE as well on the page, checks both TABLE and FIGURE matches, if the orders of table are the same and FIGURE is the table then disregard
    
    # Open the PDF and get the specific page
    doc = pymupdf.open(milstdpdf_file_path)
    page_num = table.page - 1  # Convert to 0-index
    page = doc.load_page(page_num)

    # Extract text blocks from the page
    text_blocks = page.get_text("blocks")

    # Filter blocks that start with "TABLE" in all caps and remove slash characters
    table_titles = [
        block[4].replace('\n', '').replace('/', '')  # Remove newline and slash characters
        for block in text_blocks
        if block[4].strip().startswith("TABLE")
    ]

    # Use the table's order to select the corresponding title
    # Assuming table.order is 1-based index and matches the order of appearance in the PDF
    table_order = table.order - 1  # Convert to 0-based index for Python lists
    if 0 <= table_order < len(table_titles):
        title = table_titles[table_order]
    else:
        title = "No Title Found"

    return title
    
def main():
    extract_tables()
    return

if __name__ == "__main__":
    main()

# TODO add figure/image extraction plus table extraction for that same page to be paired with
# TODO add api update requests to google sheets doc after each table
# TODO use a sheet storage for persistence in pages alreayd processed for fast processing