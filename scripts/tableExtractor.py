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
    return pageNumberSet


# TODO add figure/image extraction plus table extraction for that same page to be paired with
# TODO add api update requests to google sheets doc after each table
# TODO use a sheet storage for persistence in pages alreayd processed for fast processing


tables_page_numbers = extract_pageNumbers_from_file(tabledata_file_path)
# print(len(tables_page_numbers[1]))
# print(tables_page_numbers[0])

figures_page_numbers = extract_pageNumbers_from_file(figuredata_file_path)
# print(len(figures_page_numbers[1]))
# print(figures_page_numbers[0])
page_offset = 17 # page 1 starts after page 17, page offset + extracted page = correct page number

for i in range(len(tables_page_numbers)):
    corrected_page_number = tables_page_numbers[i] + page_offset
    currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(corrected_page_number))
    # while True:
    #     currTables = camelot.read_pdf(milstdpdf_file_path,pages = str(47 + page_offset))

    for j in range(len(currTables)):
        # print(currTable[i].df)
        print("Page Number: " + str(tables_page_numbers[i]) + str(currTables[j].parsing_report))
        if currTables[j].parsing_report['accuracy'] >= 97:
            break