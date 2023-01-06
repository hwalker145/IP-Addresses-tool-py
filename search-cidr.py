from netaddr import IPNetwork, IPAddress, cidr_merge, all_matching_cidrs
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pathlib import Path
import sys

# checks if there is the correct amount of arguments on the CLI
if len(sys.argv) < 3:
    print("USAGE: \nbook.xlsx address")
    exit()

# checks the xlsx file in the CLI if it is really a .xlsx
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

def findHeader(sheet: Worksheet, name: str) -> int:
    for cell in sheet[1]:
        if cell.value == name:
            return cell.column
        if not cell.value:
            return 0

def readSheet(sheet: Worksheet, ranges: list[IPNetwork]):
    # finds the column with the relevant data
    subnet_column: int = findHeader(sheet, "Subnet")
    # iterates through that column and adds IPNetwork objects into the list
    for column in sheet.iter_cols(subnet_column, subnet_column):
        for cell in column[1:]:
            # only adds to the list if it returns true (is not zero)
            if cell.value:
                # I call the stringtorange function to make sure each address
                # is a proper cidr address/range
                ranges.append(stringToRange(cell.value))
    # returns final product
    return ranges

address = sys.argv[2]

wb: Workbook = load_workbook(sys.argv[1])

ranges: list[IPNetwork] = []

# calling the readSheet subroutine
for ws in wb.worksheets:
    readSheet(ws, ranges)

# calls the all_matching_cidrs function from netaddr
matches: list[IPNetwork] = all_matching_cidrs(address, ranges)

# reports
print(str(len(matches)) + " matching cidr\n for the given address.\n")
for match in matches:
    print(match.cidr)