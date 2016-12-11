#!/usr/bin/python

import csv
import sys
import time
import re
import xlsxwriter

workbook = xlsxwriter.Workbook('Inventory2Pivot.xlsx')
worksheet = workbook.add_worksheet()

with open((sys.argv[1]), 'r') as infile:
	reader = csv.reader(infile)
	next(reader, None)  # skip the headers
	xrow = 0
	xcol = 0
	worksheet.write(xrow, xcol, "Bldg")
	worksheet.write(xrow, xcol + 1, "PortCount")
	xrow = 1
	for row in reader:
		regarray = {"regex":"ar-..-(.+)-.+-mg.+", "oregex":"ar-..-(.+)-.+-op.+", "PRregex":"ar-..-(.+)-.+-pr.+"}
		Is48 = r'.+(48).+'
		for n in regarray:
			pmatch = re.search(regarray[n], row[0])
			if re.search(regarray[n], row[0]):
				match = re.search(regarray[n], row[0])
				Bldg = match.group(1)
				if re.search(Is48, row[4]):
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
