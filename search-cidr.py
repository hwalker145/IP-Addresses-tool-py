from netaddr import IPNetwork, IPAddress, cidr_merge
# from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook, cell
from pathlib import Path
import sys

if len(sys.argv) < 3:
    print("USAGE: \nbook.xlsx address")
    exit()

if Path(sys.argv[1]).suffix != ".xlsx":
    print("USAGE: \nbook.xlsx address")
    exit()

# this function checks if the IP ranges are valid before constructing as IPNetwork()
def stringToRange(input: str) -> IPNetwork:
    # if the string contains neither . or :, then it's not an IP Range
    if(input.find('.') == -1 and input.find(':') == -1):
        return 0
    else:
        return IPNetwork(input)

address = sys.argv[2]

wb = load_workbook(sys.argv[1])

ranges: list[IPNetwork] = []

for sheetname in wb.sheetnames:
    subNetColumn = 1

    while wb[sheetname].cell(1, subNetColumn).value != "Subnet":
        subNetColumn += 1
        if not wb[sheetname].cell(1, subNetColumn).value:
            print("Could not find Subnet column in worksheet: " + sheetname)
            subNetColumn = 0

    if not subNetColumn:
        continue

    for row in range(2, wb[sheetname].max_row):
        print(wb[sheetname].cell(row, subNetColumn).value)
        
print(ranges)