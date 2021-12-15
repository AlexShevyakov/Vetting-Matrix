// SHEET: FS_sire
// Function to mark dates which are older than 1 year from TODAY()
// This modules DOES NOT work and is not being currently used

function formatFS_expired() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FSSheet = ss.getSheetByName("FS_sire");
  var range = FSSheet.getRange("N2:BW");
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
  
  .whenDateBefore(new Date('=TODAY()-365'))
  .setFontColor("#FF0000")
  .setBackground("#ffffff")
  .setRanges([range])
  .build();
  var rules = FSSheet.getConditionalFormatRules();
  rules.push(rule);
  FSSheet.setConditionalFormatRules(rules);
};
