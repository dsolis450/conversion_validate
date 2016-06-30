#Helper file that holds the header-function dictionary and the apply function

HEADER_FN = {
  'ExtPropertyID': val_propertyid,
  'AssetClass': val_assetclass,
  'AssetRank': val_assetrank,
  'AssetStatus': val_assetstatus,
  'AssetKeyword': val_assetkeyword
}

def apply(valFn, header, ows, nws, range_expr, cur):
  return valFn(header, ows, nws, range_expr, cur)
