#!/usr/bin/python

from xlrd import open_workbook,XL_CELL_TEXT,XL_CELL_DATE,XL_CELL_NUMBER,xldate_as_tuple
import csv
import datetime

rows = [[]]

def readsrc(src):

	sheet = open_workbook(src).sheet_by_index(0)
	for row_x in range(2,sheet.nrows):
		vals = []
		for col_x in range(0,sheet.ncols):
			cell = sheet.cell(row_x,col_x)
			if cell.ctype ==  XL_CELL_NUMBER:
				tr = str(cell.value).strip(".0")
				vals.append(tr)
			elif cell.ctype ==  XL_CELL_DATE:
				datetime_val = datetime.datetime(*xldate_as_tuple(cell.value,0))
				vals.append (str(datetime_val))
			elif cell.ctype ==  XL_CELL_TEXT:
				vals.append(cell.value)
			else:
				vals.append(" ")

	return vals

rows.append(readsrc("merged_2014-07-31-16-09-30.xls"));
print rows,