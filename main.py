import json
from openpyxl import load_workbook
from openpyxl.cell.read_only import ReadOnlyCell
wb2 = load_workbook('NF-SA Template 160519.xlsx', read_only=True)
print(wb2.sheetnames)
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