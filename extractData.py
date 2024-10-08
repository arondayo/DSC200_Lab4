# This function takes in the workbook and the section of data, and returns a dictionary with the filtered data
def extract_data(workbook_obj, section_str: str) -> list:
    data_dict = {}  # create dict to hold data

    rows = []  # create list of lists to hold rows while looping

    # loop through each row
    for row in workbook_obj.active[section_str]:
        row_data = []  # create list to hold cell values while looping
        str_count = 0  # counts the number of strings in each row
        contain_dash = False  # boolean to keep track if the row has a dash

        # loop through each cell in the current row
        for cell in row:

            # If the cell has a dash make note
            if cell.value == "–":
                row_data.append("-")


            # If the value is a string (The country name) increment string count so other strings will not be accepted
            elif isinstance(cell.value, str):
                if str_count == 0:
                    str_count += 1
                    row_data.append(cell.value)
            # If the value is a double or int, round and append it
            elif isinstance(cell.value, float) or isinstance(cell.value, int):
                rounded_value = int(round(cell.value))
                row_data.append(rounded_value)
        # fixes a bug where Andorra was one dash short
        if row_data[0] == "Andorra":
            row_data.append("-")
        rows.append(row_data)

    # return the cleaned data
    return rows
