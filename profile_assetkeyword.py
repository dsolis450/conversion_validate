def profile_assetkeyword(header_name, sheet, range_expr, cursor):
	cursor.execute("SELECT ask_asset_keyword_name from asm_asset_keyword ORDER BY 1;")  
	row = cursor.fetchone()
	qdata = [] 
	while row:
		qdata.append(row[0]) 
		row = cursor.fetchone()
      
	row_number = 1
	result = []
	#For cols that aren't required but will have lookup values, I want to copy all of the would be data to a list and then see if the list is empty. If empty move to next col, else process it.
	for row in sheet.iter_rows(range_string=range_expr):
		error_msg = ""
		row_number += 1
		for cell in row:
			if cell.value not in qdata:
				error_msg += "Asset keyword not found; "
  		
			#create row if errors found
			if error_msg != "":
				result.append([cell.value, row_number, error_msg])
	return result