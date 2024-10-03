import openpyxl
from extract_headers import extract_headers

filename = "data/Lab4Data.xlsx"
workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)

# range of headers
header_section = "B5:AF7"
headers_dict = extract_headers(workbook, header_section)
print(headers_dict)

# range of data
data_section = "B15:AF223"
