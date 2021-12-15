// SHEET: RejectHold
// Function to mark closed insepciton as green
//
function formatClosed() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rejectionsSheet = ss.getSheetByName("RejectHold");
  var range = rejectionsSheet.getRange("A2:Q1000");
  rejectionsSheet.clearConditionalFormatRules(); // MIND this line - it clears ALL RULES on the sheet!!!
  
  var rule = SpreadsheetApp.newConditionalFormatRule()
  
    .whenFormulaSatisfied('=$L2="Closed"')
    .setBackground("#B7E1CD")
    .setRanges([range])
    .build();
  var rules = rejectionsSheet.getConditionalFormatRules();
  rules.push(rule);
  rejectionsSheet.setConditionalFormatRules(rules);
};