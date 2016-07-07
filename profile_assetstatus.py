def profile_assetstatus(header_name, sheet, range_expr, cursor):
	cursor.execute("SELECT asg_asset_status_description from asm_assetstatus ORDER BY 1;")  
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
			if cell.value == None:
				error_msg += "Asset status required; "
			else:
				if cell.value not in qdata:
					error_msg += "Asset status not found; "
  		
			#create row if errors found
			if error_msg != "":
				result.append([cell.value, row_number, error_msg])
	return result