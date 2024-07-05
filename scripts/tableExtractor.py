import re
from sortedcontainers import SortedSet


tabledata_file_path = '../resources/tableData.txt'

# Function to extract integers from each line in a file
def extract_pageNumbers_from_file(txt_file_path):
    pageNumberList = SortedSet()

    with open(txt_file_path, 'r') as file:
        for line in file:
            # Find all integers in the line using regex
            pageNumber = re.search(r'\d+', line)
            pageNumberList.add(int(pageNumber.group()))
            # print(f"Page Number in line: {pageNumber}")
    return pageNumberList
