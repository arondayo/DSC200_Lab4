import openpyxl
import csv
from extractHeaders import extract_headers
from extractData import extract_data

filename = "data/Lab4Data.xlsx"
workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)

# range of headers
header_section = "B5:AF7"
headers_dict = extract_headers(workbook, header_section)
# print(headers_dict)

# range of data
data_section = "B15:AF211"
data_dict = extract_data(workbook, data_section)

# Write data dict and header dict to a csv file
with open("data/andersona26_weils3.csv", "w") as outfile:
    csvWriter = csv.writer(outfile, lineterminator='\n')
    csvWriter.writerow(headers_dict.values())
    count = 0
    for value in data_dict.values():
        csvWriter.writerow(value)
        count += 1

# print the number of rows outputted to the csv
print("There are {} rows in the csv!".format(count))

