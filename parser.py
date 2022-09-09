#!/usr/bin/env python3

# simple python program to parse through excel files and find duplicates
# code by Joshua Bisanda (cyb3rtooth)
# there is always room for improvement

import openpyxl
import argparse

# load the excel file
parser = argparse.ArgumentParser() # create a new parser object to grab arguments
parser.add_argument('--file', type=str, required=True, help='The excel file name to be processed')
parser.add_argument('--col', type=int, required=True, help='The column number to be processed; 0 is the first column')
args = parser.parse_args()

workbook = openpyxl.load_workbook(args.file)
# grab the active worksheet
sheet = workbook.active

# create an empty list of items to be compared
items = []

# grab all the relevant values to compare
print('[+] Grabbing values from the excel file...')
for cellObj in list(sheet.columns)[args.col]:
    items.append(cellObj.value)

# check for duplicates in the items list
newItemsList = []
listForDupes = []
print('[+] Done!')

print('[+] Checking for duplicates...')
for x in items:
    if x not in newItemsList:
        newItemsList.append(x)
    else:
        listForDupes.append(x)

# verify that the list for duplicates is not empty:
if (len(listForDupes) <= 0):
    print('[!] No duplicates were found.')
else:
    print('[+] Some duplicates were found! \n')
    for item in listForDupes:
        print(item)
