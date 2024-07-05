import re
from sortedcontainers import SortedSet
import camelot


milstdpdf_file_path = '../resources/MIL-STD-1472H.pdf'
tabledata_file_path = '../resources/tableData.txt'
figuredata_file_path = '../resources/figureData.txt'

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

# extract_titles_from_file(figuredata_file_path)
# extract_titles_from_file(tabledata_file_path)

# TODO add figure/image extraction plus table extraction for that same page to be paired with
# TODO add api update requests to google sheets doc after each table
# TODO use a sheet storage for persistence in pages alreayd processed for fast processing


tables_page_numbers = extract_pageNumbers_from_file(tabledata_file_path)

figures_page_numbers = extract_pageNumbers_from_file(figuredata_file_path)

page_offset = 17 # page 1 starts after page 17, page offset + extracted page = correct page number

for i in range(len(tables_page_numbers)):
    corrected_page_number = tables_page_numbers[i] + page_offset
    currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(corrected_page_number))
    # while True:
    #     currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(47 + page_offset))

    for j in range(len(currTables)):
        # print(currTable[i].df)
        print("Page Number: " + str(tables_page_numbers[i]) + str(currTables[j].parsing_report))
        currTables[j].to_csv('asd.csv')
        # if currTables[j].parsing_report['accuracy'] >= 97:
        #     break