
def profile_propertyid(header_name, sheet, nws, range_expr, cursor):
  cursor.execute("SELECT lgp_external_property_id FROM lgb_property ORDER BY 1;")  
  row = cursor.fetchone()
  results = []  
  while row:  
      results.append(row[0])    
      row = cursor.fetchone()
      
  row_number = 1
  error_msg = ""
  for row in sheet.iter_rows(range_string=range_expr):
      row_number += 1
      for cell in row:
  		if cell.value == None:
  			error_msg += "ExtPropertyID required; "
  		else:
  		    if cell.value not in results:
  				error_msg += "Property ID not found; "
  		
  		if error_msg != "":
  			nws.append([cell.value, row_number, error_msg])
