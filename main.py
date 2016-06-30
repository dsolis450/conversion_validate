import sys
import pymssql  
import openpyxl

from _profile_helper import HEADER_FN
from _profile_helper import apply
from Tkinter import Tk
from tkFileDialog import askopenfilename
from openpyxl.styles import Color, PatternFill, Font, Border


database_in = 'SUPPORT' + raw_input('Enter support database number: ')
try:
    conn = pymssql.connect(server='ATLASSPSQL1', user='sa', password='titp4sa', database=database_in)  
except pymssql.Error as e:
	print "An error has occurred: ", e
	sys.exit(1)
cursor = conn.cursor()  
Tk().withdraw()
filename = askopenfilename()

wb = openpyxl.load_workbook(filename, use_iterators=True)
sheet = wb.get_sheet_by_name('Assets')
nwb = openpyxl.Workbook(write_only=True)

row_count = sheet.max_row

#header_lst = (openpyxl command to get first row cells)
for header in header_lst:
	if header.value not in HEADER_FUNC_DICT:
		pass
	else:
		nws = nwb.create_sheet(header.value)
		nws.append([header.value,'Row','Errors'])
		#header_col = (openpyxl command to get the col of header) header.col
		range_expr = header_col + '2:' + header_col + str(row_count)
		apply(HEADER_FN[header.value], header.value, sheet, nws, range_expr, cursor)