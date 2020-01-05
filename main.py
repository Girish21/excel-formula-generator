import re
from openpyxl import load_workbook, Workbook
from openpyxl.cell.read_only import ReadOnlyCell

# regex to capture patterns like 'Ace-SP&L'!C20*, 'Ace-SP&L'!C$120, C356, D$89+
root_pattern_regex = r"('(?:(?:Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(?:(?:[0-9])+))|(?:[A-Z]\$?(?:(?:[0-9])+)))(\+|\-|\*|\/)?"
# regex to extract (Ace-SP&L, 20) from 'Ace-SP&L'!C20, 439 from D$439
extract_regex = r"(?:(?:'(Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(\d{1,}))|[A-Z]\$?(\d{1,}))"


"""
    method to recursively find the name of the row for the given excel coordinate
    @params
        extracted_expression - excel coordinate (eg: C40, F$567, 'Ace-SP&L'!C20)
        root_dict - ref to root dict to fetch row name
"""


def find_name(extracted_expression, root_dict):
    if len(list(
            filter(None, list(re.findall(extract_regex, extracted_expression))))) == 0:
        return extracted_expression
    extracted_expression = list(
        filter(None, list(re.findall(extract_regex, extracted_expression)[0])))
    if len(extracted_expression) > 1:
        if int(extracted_expression[1]) in root_dict[extracted_expression[0]]:
            return find_name(root_dict[extracted_expression[0]][int(extracted_expression[1])], root_dict)
        else:
            return extracted_expression[1]
    else:
        if int(extracted_expression[0]) in root_dict['SA-Ratios']:
            return find_name(root_dict['SA-Ratios'][int(extracted_expression[0])], root_dict)
        else:
            return extracted_expression[0]


"""
    method to build the target parsed excel
    @params
        book - ref to the loaded excel book
        root_dict - ref to root dict
        dest_sheet_ref - ref to the target excel to write to
"""


def build_excel(book, root_dict, dest_sheet_ref):
    row = 1
    for cells in book.iter_rows(min_col=2, max_col=19, min_row=7):
        for cell in cells:
            if type(cell) == ReadOnlyCell and cell.value != None and re.match(r'^-?\d+(?:\.\d+)?$', str(cell.value)) is None:
                if cell.column != 2 and len(list(re.findall(root_pattern_regex, cell.value))) > 0:
                    contents = list(filter(None, re.split(
                        root_pattern_regex, cell.value)))
                    parsed_first_part = list(filter(None, re.findall(
                        r"=?(SUM\(|AVERAGE\(|\()", contents[0])))
                    if len(parsed_first_part) > 0:
                        contents[0] = parsed_first_part[0]
                    else:
                        contents.pop(0)
                    for i in range(len(contents)):
                        contents[i] = find_name(contents[i], root_dict)
                    dest_sheet_ref["{}{}".format(
                        "B", str(row))] = ' '.join(contents)
                    row += 1
                    break
                else:
                    if len(list(re.findall(root_pattern_regex, cell.value))) > 0:
                        contents = list(filter(None, re.split(
                            root_pattern_regex, cell.value)))
                        parsed_first_part = list(filter(None, re.findall(
                            r"=(SUM\(|AVERAGE\(|\()", contents[0])))
                        if len(parsed_first_part) > 0:
                            contents[0] = parsed_first_part[0]
                        else:
                            contents.pop(0)
                        for i in range(len(contents)):
                            contents[i] = find_name(contents[i], root_dict)
                        dest_sheet_ref["A{}".format(
                            str(row))] = ' '.join(contents)
                    else:
                        dest_sheet_ref["A{}".format(str(row))] = cell.value


""" 
    method to generate dict for faster lookup
    @params
        book - ref to the loaded excel book
        root_dict - ref to root dict
        sheet_list - list of sheets to parse and generate lookup
        min_col
        max_col
        min_row
"""


def generate_dict(book, root_dict, sheet_list, min_col, max_col, min_row):
    for sheet in sheet_list:
        temp_dict = {}  # {[key: number]: string}
        for cell in book[sheet].iter_rows(min_col=min_col, max_col=max_col, min_row=min_row):
            # check to discard None values and empty cells
            if type(cell[0]) == ReadOnlyCell and cell[0].value != None:
                # remove unicode value \u00A0
                temp_dict[cell[0].row] = cell[0].value.strip()
        root_dict[sheet] = temp_dict


def main():
    # Open workbook to write to
    write_book = Workbook()
    active_write_book = write_book.active

    # Open the given file to parse
    read_book = load_workbook('NF-SA Template 160519.xlsx', read_only=True)

    # List of helper sheets needs to be parsed
    sheets = ['Ace-SP&L', 'Ace-SBS', 'Ace-SCFS']
    # Root dict to hold the parsed data
    root_dict = {}

    # generate the dict for lookup of row names
    generate_dict(read_book, root_dict, ['SA-Ratios'], 2, 2, 7)
    generate_dict(read_book, root_dict, sheets, 1, 1, 4)

    # build the target excel
    build_excel(read_book['SA-Ratios'], root_dict, active_write_book)

    write_book.save(filename="result.xlsx")

    pass


if __name__ == "__main__":
    main()
