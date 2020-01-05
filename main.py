import json
import re
from openpyxl import load_workbook, Workbook
from openpyxl.cell.read_only import ReadOnlyCell
write_book = Workbook()
wb = write_book.active
wb2 = load_workbook('NF-SA Template 160519.xlsx', read_only=True)
sa_ratio = wb2['SA-Ratios']
helper_sheets = ['Ace-SP&L', 'Ace-SBS', 'Ace-SCFS']
max_row = sa_ratio.max_row

row_dict = {}

temp_dict = {}
for cell in sa_ratio.iter_rows(min_col=2, max_col=2, min_row=7):
    if type(cell[0]) == ReadOnlyCell and cell[0].value != None:
        temp_dict[cell[0].row] = cell[0].value.strip()

row_dict['SA-Ratios'] = temp_dict

for sheet in helper_sheets:
    sheet_cols = {}
    for cell in wb2[sheet].iter_rows(min_col=1, max_col=1, min_row=4):
        if type(cell[0]) == ReadOnlyCell and cell[0].value != None:
            sheet_cols[cell[0].row] = cell[0].value.strip()
    row_dict[sheet] = sheet_cols

root_pattern_regex = r"('(?:(?:Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(?:(?:[0-9])+))|(?:[A-Z]\$?(?:(?:[0-9])+)))(\+|\-|\*|\/)?"
extract_regex = r"(?:(?:'(Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(\d{1,}))|[A-Z]\$?(\d{1,}))"
# i = 0
# for row in sa_ratio.rows:
#     canIterate = False
#     for cell in row:
#         if type(cell) == ReadOnlyCell and cell.value != None:
#             if cell.column == '2' and cell.value != None:
#                 canIterate = True
#             if  len(list(re.findall(root_pattern_regex, cell.value))) > 0:
#                 contents = list(filter(None, re.split(
#                     root_pattern_regex, cell.value)))
#                 for expression in contents:
#                     if re.match(root_pattern_regex, expression):
#                         print(cell.coordinate, expression)
#                 break
#             else:
#                 wb[cell.coordinate] = cell.value
#     i += 1
#     if i == 10:
#         break

for cells in sa_ratio.iter_rows(min_col=2, max_col=19):
    for cell in cells:
        if type(cell) == ReadOnlyCell and cell.value != None and re.match(r'^-?\d+(?:\.\d+)?$', str(cell.value)) is None:
            if len(list(re.findall(root_pattern_regex, cell.value))) > 0:
                # print(list(filter(None, re.split(
                #     root_pattern_regex, cell.value))))
                contents = list(filter(None, re.split(
                    root_pattern_regex, cell.value)))
                contents.pop(0)
                for i in range(len(contents)):
                    expression = contents[i]
                    if len(list(re.findall(extract_regex, expression))) > 0:
                        extracted_expression = list(
                            filter(None, list(re.findall(extract_regex, expression)[0])))
                        if (len(extracted_expression) > 1):
                            if int(
                                    extracted_expression[1]) in row_dict[extracted_expression[0]]:
                                contents[i] = row_dict[extracted_expression[0]][int(
                                    extracted_expression[1])]
                        else:
                            if int(
                                    extracted_expression[0]) in row_dict['SA-Ratios']:
                                contents[i] = row_dict['SA-Ratios'][int(
                                    extracted_expression[0])]
                wb["{}{}".format(
                    "C", str(cell.row))] = ' '.join(contents)
                break
            else:
                wb[cell.coordinate] = cell.value

write_book.save(filename="test.xlsx")

# reg=r"'(?:Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(?:(?:[0-9])+)(\+|\-|\*|\/)"
# print(re.findall(reg, "'Ace-SP&L'!C20/'Ace-SP&L'!C$6*'Ace-SP&L'!C$6"))


# (?:'(?:(?:Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(?:(?:[0-9])+))|(?:[A-Z](?:(?:[0-9])+)))(\+|\-|\*|\/)
# print(re.findall(reg, "(S212/I212)^(1/10)-1"))
# ['/']

# ('(?:(?:Ace-(?:SP&L|SBS|SCFS))'![A-Z]\$?(?:(?:[0-9])+))|(?:[A-Z]\$?(?:(?:[0-9])+)))(\+|\-|\*|\/)?
# print(re.findall(reg, "'Ace-SP&L'!C20/'Ace-SP&L'!C$6*'Ace-SP&L'!C$6"))
# ['/', '*', '']

# =C8-C10
# =C9/C$8
# =B890
# ='Ace-SP&L'!C16
# ='Ace-SP&L'!C17+'Ace-SP&L'!C6
# =100%-C16
# ='Ace-SP&L'!C20/'Ace-SP&L'!C$6*'Ace-SP&L'!C$6
# =C30/AVERAGE(D31)
# =C92+C93+C94
# =S99/R99-1
# =(S69/I69)^(1/10)-1
# =S166/AVERAGE(R167:S167)
# =(S212/I212)^(1/10)-1
# =S365-'Ace-SP&L'!R203
