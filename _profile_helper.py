#Helper file that holds the header-function dictionary and the apply function
from profile_assetclass import profile_assetclass
from profile_propertyid import profile_propertyid
from profile_assetname import profile_assetname
from profile_assetnumber import profile_assetnumber
from profile_serialnumber import profile_serialnumber
from profile_assetrank import profile_assetrank

HEADER_FN = {
	'PROPERTYID': profile_propertyid,
	'ASSETCLASS': profile_assetclass,
	'ASSETNAME': profile_assetname,
	'ASSETNUMBER': profile_assetnumber,
	'SERIALNUMBER': profile_serialnumber,
	'ASSETRANK': profile_assetrank
 # 'MAKE': profile_make,
 # 'MODEL': profile_model,
 # 'ASSETSTATUS': profile_assetstatus,
 # 'ASSETKEYWORD': profile_assetkeyword
}

def apply(valFn, header, ows, range_expr, cur):
  return valFn(header, ows, range_expr, cur)
