#!/usr/bin/env python3

import argparse
from openpyxl import load_workbook
import json

parser = argparse.ArgumentParser(description='Convert xlsx to json')
parser.add_argument('file', metavar='F', type=str,
                            help='input file')
parser.add_argument('--header_row', type=int, default=1, help='header row')
parser.add_argument('--first_data_row', type=int, default=2, help='first data row')

args = parser.parse_args()

wb = load_workbook(args.file)
ws = wb.active

header = []
for j in range(1, ws.max_column + 1):
    header.append(ws.cell(row=args.header_row, column=j).value)

data = []
rows = ws.rows
for i in range(args.first_data_row - 1, ws.max_row):
    row = rows[i]
    datarow = {}
    for j in range(0, len(row)):
        datarow[header[j]] = row[j].value
    data.append(datarow)

outdata = {'header': header, 'data': data}

print(json.dumps(outdata, ensure_ascii=False))
