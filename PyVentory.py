#!/usr/bin/python

import csv
import sys
import time
import re
import xlsxwriter

workbook = xlsxwriter.Workbook('Inventory2Pivot.xlsx')
worksheet = workbook.add_worksheet()

with open((sys.argv[1]), 'r') as infile, open((sys.argv[2]), 'w') as outfile:
    reader = csv.reader(infile)
    next(reader, None)  # skip the headers
    xrow = 0
    xcol = 0
    worksheet.write(xrow, xcol, "Bldg")
    worksheet.write(xrow, xcol + 1, "PortCount")
    xrow = 1
    for row in reader:
		regex = r'..-..-(.+)-.+-mg.+'
		pregex = r'.+(48).+'
		oregex = r'..-..-(.+)-.+-op.+'
		PRregex = r'..-..-(.+)-.+-pr.+'
		if re.search(regex, row[0]):
			match = re.search(regex, row[0])
			Bldg = match.group(1)
			if re.search(pregex, row[4]):
				prt = "Medium"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)
			else:
				prt = "Small"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)
		elif re.search(oregex, row[0]):
			match = re.search(oregex, row[0])
			Bldg = match.group(1)
			if re.search(pregex, row[4]):
				prt = "Medium"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)
			else:
				prt = "Small"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)
		elif re.search(PRregex, row[0]):
			match = re.search(PRregex, row[0])
			Bldg = match.group(1)
			if re.search(pregex, row[4]):
				prt = "Medium"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)
			else:
				prt = "Small"
				worksheet.write(xrow, xcol, Bldg)
				worksheet.write(xrow, xcol + 1, prt)			
		xrow += 1
workbook.close()
print "All Done"
