import re
import os
import csv
from definitionsScraper import exponential_backoff

import camelot
import pymupdf

import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials

from sortedcontainers import SortedSet

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
def extract_pageNumbers_from_file(txt_file_path):
    pageNumberSet = SortedSet()

    with open(txt_file_path, 'r') as file:
        for line in file:
            pageNumbers = re.findall(r'\d+', line)
            pageNumberSet.add(int(pageNumbers[-1]))

    return pageNumberSet

def upload_csv_file(csv_file_path):
    # TODO add center of table titles AND remove file numbers in the front
    # Extract the name of the CSV file without the extension
    full_sheet_name = os.path.splitext(os.path.basename(csv_file_path))[0]

    file_num_rmv_csv_title = full_sheet_name.split(" ")
    file_num_rmv_csv_title.pop(0)
    file_num_rmv_csv_title = ' '.join(file_num_rmv_csv_title)

    new_sheet_name = convert_to_sheet_name(full_sheet_name).strip()

    # Read data from the CSV file
    with open(csv_file_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        data = list(reader)

    num_rows = len(data) #inserts a row so dont need to set the size to exact
    num_cols = max(len(row) for row in data)

    worksheet = None
    
    end_col_letter = chr(ord('A') + num_cols - 1)  # Convert column number to letter for the last column
    if "Continued" in new_sheet_name:

        existing_sheet_name = new_sheet_name.replace(" Continued", "").strip()
        request_attempt = 0
        while True:
            try:
                worksheet = usage_limit_retry(lambda: cleanDataFile.worksheet(existing_sheet_name))
                break
            except gspread.exceptions.WorksheetNotFound:
                print(f"Retrying {existing_sheet_name} look up...")
                request_attempt += 1
                exponential_backoff(request_attempt)

        continued_title_index = worksheet.row_count + 1
        usage_limit_retry(lambda: worksheet.resize(rows=continued_title_index, cols=num_cols)) # resize sheet to be exact
        merge_range = f"A{worksheet.row_count}:{end_col_letter}{continued_title_index}" # row_count to use the index of the last row in the sheet for continuation
        usage_limit_retry(lambda: worksheet.merge_cells(merge_range))
        usage_limit_retry(lambda: worksheet.update([[file_num_rmv_csv_title]], f"A{continued_title_index}"))
        format_cell_range(cleanDataFile.worksheet(existing_sheet_name), merge_range, CellFormat(horizontalAlignment='CENTER'))
        # Write the data to the new Google Sheet
        for row in data:
            usage_limit_retry(lambda: worksheet.append_row(row))

    else:
        try:
            worksheet = usage_limit_retry(lambda: cleanDataFile.worksheet(new_sheet_name))
            print(f"Sheet '{new_sheet_name}' already exists.")
        except gspread.exceptions.WorksheetNotFound:
            # Create a new sheet with the name derived from the CSV file name
            worksheet = usage_limit_retry(lambda: cleanDataFile.add_worksheet(title=new_sheet_name, rows=num_rows, cols=num_cols))
            print(f"New sheet '{new_sheet_name}' created.")

        merge_range = f"A1:{end_col_letter}1"
        usage_limit_retry(lambda: worksheet.merge_cells(merge_range))
        usage_limit_retry(lambda: worksheet.update([[file_num_rmv_csv_title]], f"A1"))
        format_cell_range(cleanDataFile.worksheet(new_sheet_name), merge_range, CellFormat(horizontalAlignment='CENTER'))

        # Write the data to the new Google Sheet
        for row_index, row in enumerate(data, start=2):
            usage_limit_retry(lambda: worksheet.insert_row(row, row_index))

        usage_limit_retry(lambda: worksheet.resize(rows=num_rows, cols=num_cols)) # resize sheet to be exact
    print(f"Data written to sheet '{new_sheet_name}' successfully.")

def extract_tables(tables_page_numbers: SortedSet):

    for curr_page_number in tables_page_numbers:
        corrected_page_number = curr_page_number + page_offset

        if table_titles_matches(corrected_page_number + 1): # checks if the next page has a TABLE title and adds that page number into the sorted set, if it doesn't that means there isn't a TABLE continued. helps shorten time to process 
            tables_page_numbers.add(curr_page_number + 1)

        currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(corrected_page_number))

        for currTable in currTables:
            print("Page Number: " + str(curr_page_number) + str(currTable.parsing_report))

            currTable_title = corresponding_table_title_extraction(currTable)
            
            if currTable_title == None:
                continue
            
            currTable_filepath = os.path.join(saved_tables_csv_filepath, f"{currTable_title}.csv")
            currTable.to_csv(currTable_filepath)
            # upload_csv_file(currTable_filepath)

def extract_page_numbers(start, end): # fixed missing page numbers for tables in between by using range of start and end tables number
    return list(range(start, end + 1))

def corresponding_table_title_extraction(table):
    table_titles = table_titles_matches(table.page)

    # Use the table's order to select the corresponding title
    # Assuming table.order is 1-based index and matches the order of appearance in the PDF
    table_order = table.order - 1  # Convert to 0-based index for Python lists
    file_number = len([name for name in os.listdir(saved_tables_csv_filepath) if os.path.isfile(os.path.join(saved_tables_csv_filepath, name))]) + 1 

    if 0 <= table_order < len(table_titles):
        title = str(file_number) + ". " + re.sub(r"[\s\.0-9]+$", "", table_titles[table_order]) # removes trailing: periods, spaces, numbers
    else:
        title = None

    return title

def table_titles_matches(one_index_page_number):
    # Open the PDF and get the specific page
    doc = pymupdf.open(milstdpdf_file_path)
    page_num = one_index_page_number - 1  # Convert to 0-index
    page = doc.load_page(page_num)

    # Extract text blocks from the page
    text_blocks = page.get_text("blocks")

    # Filter blocks that start with "TABLE" in all caps and remove slash characters
    table_titles = [
        block[4].replace('\n', '').replace('/', '')  # Remove newline and slash characters
        for block in text_blocks
        if block[4].strip().startswith("TABLE")
    ]

    return table_titles

def convert_to_sheet_name(filepath):
    parts = filepath.split(' ')
    table_num_part = []

    table_num_part.append(parts[1]) # get only TABLE and Number
    table_num_part.append(parts[2])

    # Check if "Continued" exists in the title
    if "Continued" in parts:
        table_num_part.append(" Continued")

    table_num_title = ' '.join(table_num_part)

    return table_num_title

def main():
    # extract_tables(extract_pageNumbers_from_file(tabledata_file_path)) # full document, * need to double checkfile names, doesn't always find the tables
    extract_tables(SortedSet([419])) # specific pages to search for tables
    # upload_csv_file()
    return

if __name__ == "__main__":
    main()

# TODO add figure/image extraction plus table extraction for that same page to be paired with
# TODO strip each text added with broken characters
# TODO possibly add retry of pages until a table is found