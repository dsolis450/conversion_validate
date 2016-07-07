def profile_'''colname'''(header_name, sheet, range_expr, cursor):
	cursor.execute('''"query"''')  
	row = cursor.fetchone()
	qdata = [] 
	while row:
		qdata.append(row[0]) 
		row = cursor.fetchone()
      
	row_number = 1
	result = []
	for row in sheet.iter_rows(range_string=range_expr):
		error_msg = ""
		row_number += 1
		for cell in row:
			''' Uncomment if required value
			if cell.value == None:
				error_msg += "FIELD required; "
			else:
			'''
			#Indent if working with required value logic above
			if cell.value  '''enter logic to check for error''':
				error_msg += "enter error message; "
  		
			#create row if errors found
			if error_msg != "":
				result.append([cell.value, row_number, error_msg])
	return result