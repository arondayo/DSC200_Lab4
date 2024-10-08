import openpyxl
import csv

from openpyxl.styles.builtins import total

from extractheaders import extract_headers
from extractdata import extract_data


#open the excel file
filename = "data/Lab4Data.xlsx"
workbook = openpyxl.load_workbook(filename, read_only=True, data_only=True)

# range of headers
header_section = "B5:AF7"
headers_dict = extract_headers(workbook, header_section)


# range of data
data_section = "B15:AF211"
data_list = extract_data(workbook, data_section)

output_headers = ["CountryName", "CategoryName","CategoryTotal"]

#This section takes the data from header dictionary and Data list and formats them correctly in a list of lists to be written to a csv
#Creates an empty list for the output to be put into
output_list = []
for data in data_list: #Loop through each country's data from datalist

    row = []
    country = data[0]
    row.append(country) #set the country name for the CountryName column
    i = 1 #Keeps track of which data element the loop is on
    for key, value in headers_dict.items(): #Loops through each header for each country
        if value == "Countries and areas":
            continue
        row.append(value) #Add in the value for the CategoryName column
        if i < len(data):  # Ensure i does not go out of bounds
            if data[i] == "-" or data[i] == 0: #Check if the row is equal to zero or a -, throw out if true
                i+=1
                row = [country] #Reset the row list to be ready for the next iteration
                continue
            else: #add in the value for the CategoryTotal Column
                row.append(data[i])
        i += 1
        output_list.append(row) #Add the row to the output list
        row = [country] #Reset the row list to be ready for the next iteration





#Write data dict and header dict to a csv file
with open("weils3.csv", "w") as outfile:
    csvWriter = csv.writer(outfile, lineterminator='\n')
    csvWriter.writerow(output_headers)
    count = 0
    for row in output_list:
        csvWriter.writerow(row)
        count += 1

#print the number of rows outputted to the csv
print("There are {} rows in the csv!".format(count))
