def profile_assetrank(header_name, sheet, range_expr, cursor):
	cursor.execute("SELECT asr_asset_rank_description FROM asm_assetrank ORDER BY 1;")  
	row = cursor.fetchone()
	results = []  
	while row:  
		results.append(row[0])    
		row = cursor.fetchone()
      
	row_number = 1
	result = []
	for row in sheet.iter_rows(range_string=range_expr):
		error_msg = ""
		row_number += 1
		for cell in row:
			if cell.value == None:
				error_msg += "Asset rank required; "
			else:
				if cell.value not in results:
					error_msg += "Asset rank not found; "
  	
			#create row if errors found
			if error_msg != "":
				result.append([cell.value, row_number, error_msg])
	return result