def profile_assetclass(header_name, sheet, range_expr, cursor):
  cursor.execute("SELECT asc_asset_class_description FROM asm_assetclass ORDER BY 1;")  
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
  			error_msg += "AssetClass required; "
  		else:
  		    if cell.value not in results:
  				error_msg += "Asset class not found; "
  				
  		if error_msg != "": 
			result.append([cell.value, row_number, error_msg])
  return result
