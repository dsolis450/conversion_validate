"""
data structure:

Dictionary - key is column header, subdict with info (query, function) maybe just the function and the function will have the query? makes more sense
"""
headerInfo = {'ExtPropertyID' 	: 	valExtPropID,
						'AssetClass' 		: 	valAssetClass,
						'AssetRank' 			: 	valAssetRank,
						'AssetStatus' 		: 	valAssetStatus,
						'AssetKeyword' 	: 	valAssetKeyword}

def valExtPropID(cursor, nwb, sheet):
	cursor.execute("SELECT lgp_external_property_id FROM lgb_property ORDER BY 1;")  
	row = cursor.fetchone()
	results = []  
	while row:  
		results.append(row[0])    
		row = cursor.fetchone()
	nws = nwb.create_sheet('ExtPropID')
	range_expr = 'A1:A' + str(row_count)
	for row in sheet.iter_rows(range_string=range_expr):
		for cell in row:
			if cell.value == None:
				nws.append([cell.value])
			else:
				if cell.value in results:
					pass
				else:
					nws.append([cell.value])

def apply(valFunc, cur, nwb, ows):
	return valFunc(cur, nwb, ows)
					


def main():
	headers = #read first row of values in excel sheet
	for header in headers:
		apply(headerInfo[header], cursor, nwb, sheet)	
		
		
		
		
