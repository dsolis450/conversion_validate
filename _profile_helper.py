#Helper file that holds the header-function dictionary and the apply function

HEADER_FN = {
  'ExtPropertyID': val_propertyid,
  'AssetClass': val_assetclass,
  'AssetName': val_assetname,
  'AssetNumber': val_assetnumber,
  'SerialNumber': val_serialnumber,
  'AssetRank': val_assetrank,
  'Make': val_make,
  'Model': val_model,
  'AssetStatus': val_assetstatus,
  'AssetKeyword': val_assetkeyword
}

def apply(valFn, header, ows, nws, range_expr, cur):
  return valFn(header, ows, nws, range_expr, cur)
