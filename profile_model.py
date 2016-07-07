def profile_model(header_name, sheet, range_expr, cursor):  
	row_number = 1
	result = []
	for row in sheet.iter_rows(range_string=range_expr):
		error_msg = ""
		row_number += 1
		for cell in row:
			if cell.value == None:
				error_msg += "Model required; "
				
			#create row if errors found
			if error_msg != "":
				result.append([cell.value, row_number, error_msg])
	return result