import xlrd, xlwt
from xlutils.copy import copy
import os

def normalize_delimiters(staging_dir, excel_file, check_sheets, DRY_RUN):
	'''Takes an Excel file and looks for columns that have both numbers and text in them (ignoring Rows 1 & 2)
	Assumes that mixed-type columns should all be numbers and that European-style numbers were used
		(using commas to demarcate the tenths place), causing numbers to be interpreted as text.
	Converts commas to decimal points in all such rows, then resaves the same file
	''' 

	rb = xlrd.open_workbook(os.path.join(staging_dir,excel_file), formatting_info = True)
	wb = copy(rb)

	for sheet_num in range(rb.nsheets):
		rs = rb.sheet_by_index(sheet_num)
		if rs.name in check_sheets:
			ws = wb.get_sheet(sheet_num)

			#figure out last real value in sheet
			blank_rows = 0
			for row in range(2,rs.nrows):
				for col in range(1,5):
					if rs.cell(row,col).value != '':
						break
					elif col == 4 :
						blank_rows +=1
				if blank_rows >= 4:
					end_of_data = row - 3
					break
			if not end_of_data:
				end_of_data = rs.nrows

			#now check columns for things to convert
			for col in range(rs.ncols):
				if xlrd.XL_CELL_NUMBER in rs.col_types(col,2,end_of_data) and xlrd.XL_CELL_TEXT in rs.col_types(col,2,end_of_data):
					converted = False
					for row in range(rs.nrows)[2:end_of_data]:
						if rs.cell(row,col).ctype==xlrd.XL_CELL_TEXT and rs.cell(row,col).value!='':
							converted = True
							try:
								ws.write(row,col,float(rs.cell(row,col).value.replace(',','.')))
							except(ValueError):
								ws.write(row,col,rs.cell(row,col).value.replace(',','.'))
								if DRY_RUN: print("'%s' in row %s is not a number, so leaving as text. Any commas were converted" % (rs.cell(row,col).value,row))
							except:
								if DRY_RUN: print("Unable to convert value '%s' in row %s" % (rs.cell(row,col).value,row))
								pass
					if converted == True: print("Converted column '%s' in sheet '%s' to numbers" % (rs.cell(1,col).value, rs.name))

	if excel_file[-5] == '.xlsx':
		excel_file = excel_file[:-1]
	wb.save(os.path.join(staging_dir,excel_file))

	return excel_file