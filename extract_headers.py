from functools import reduce
import re, string


# Given an incoming openpyxl.workbook object and section
# produce a dict of headers, {col_str0: col_name0, col_str1: col_name1, ...}

# Algorithm:
# Cascades down the columns constructing a compound key out of the headers it finds**
# Starts a new base key when it encounters a new key at the top of the column
#   **
#   if we encounter a new middle key somewhere before the end of the column we update the key
#   but if we encounter another middle key we reconstruct our compound key
# when we hit the end of col and there's a new key entry compound the key but don't change the prev top key
# if it doesn't run into any more keys to compound, output whatever the prev top key was
# when the bottom of the col is reached add key, value to output dict


def extract_headers(workbook_obj, section_str: str) -> dict:
    def calc_col_label(range_string: str, col_index: int) -> str:
        # Borrowed from stack overflow: https://stackoverflow.com/questions/48983939/convert-a-number-to-excel-s-base-26
        # converts base 26 character strings to integers ("AA" -> 27)
        def to_excel(num):
            def divmod_excel(n):
                a, b = divmod(n, 26)
                if b == 0:
                    return a - 1, b + 26
                return a, b

            chars = []
            while num > 0:
                num, d = divmod_excel(num)
                chars.append(string.ascii_uppercase[d - 1])
            return ''.join(reversed(chars))

        # converts integers to base 26 character strings (26 -> "Z")
        def from_excel(chars):
            return reduce(lambda r, x: r * 26 + x + 1, map(string.ascii_uppercase.index, chars), 0)

        # End of Borrowing

        col = re.findall("^[A-z]+", range_string.split(":")[0])[0]
        return to_excel(from_excel(col) + col_index)

    rows = []
    for row in workbook_obj.active[section_str]:
        row_data = []
        for cell in row:
            row_data.append(cell.value)
        rows.append(row_data)

    # Cascades down the columns constructing a compound key out of the headers it finds*
    # Starts a new base key when it encounters a new key at the top of the column
    output = {}

    root_top = ""
    previous_top = ""
    for i_col in range(len(rows[0])):
        top_was_changed = True
        mid_was_changed = True
        for j_row in range(len(rows)):
            item = rows[j_row][i_col]
            if item is not None:
                key = re.sub("\\n", "", rows[j_row][i_col])
                if j_row == 0:
                    # if it's at the top of a column we set that as the new prev top key
                    top_was_changed = False
                    mid_was_changed = False
                    root_top = key
                    previous_top = key
                elif j_row != 0 and j_row != len(rows) - 1:
                    # **
                    # if we encounter a new middle key somewhere before the end of the column we update the key
                    # but if we encounter another middle key we reconstruct our compound key
                    if not mid_was_changed:
                        mid_was_changed = True
                        top_was_changed = True
                        previous_top = previous_top + "_" + key
                    else:
                        previous_top = root_top + "_" + key
                else:
                    # when we hit the end of col and there's a new key entry compound the key
                    # but don't change the prev top key
                    top_was_changed = True
                    final_key = previous_top + "_" + key
                    output[calc_col_label(section_str, i_col)] = final_key
                    # print(f" [{i_col}][{j_row}]\t|\t{final_key}")
            elif j_row == len(rows) - 1 and not top_was_changed:
                # if it doesn't run into any more keys to compound, output whatever the prev top key was
                output[calc_col_label(section_str, i_col)] = previous_top
                # print(f" [{i_col}][{j_row}]\t|\t{previous_top}")
    return output
