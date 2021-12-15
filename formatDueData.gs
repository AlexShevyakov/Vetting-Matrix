
// This module provides conditional formatting for DueData sheet

/*
function getInspectionStatus(statusCase) {
  return '=P2:P100="' + statusCase +'"';
}

*/
function applyFormattingDueData () {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dueData = ss.getSheetByName("DueData");
  let rangeDueData = dueData.getRange("A3:B59");
  let rangeComments = dueData.getRange("B3:B59");
    
   dueData.clearConditionalFormatRules(); // MIND this line - it clears ALL RULES on the sheet!!!
  
   // RULE: Comments missing
  var ruleCommentsMissing = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(A3:A326<>"",B3:B326="")')
    .setFontColor("#e06666")
    .setRanges([rangeComments])
    .build();
  var rules = dueData.getConditionalFormatRules();
  rules.push(ruleCommentsMissing);
  dueData.setConditionalFormatRules(rules);

  
  // EDOI: Acknowledged
  var ruleAcknowledged_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Acknowledged'))
    .setFontColor("#0000FF")
    .setItalic(true)
    .setRanges([rangeDateRequested])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleAcknowledged_d);
  activity.setConditionalFormatRules(rules);
  
  
 // STATUS: Scheduled
  var ruleConfirmed_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Scheduled'))
    .setFontColor("#0000FF")
    .setItalic(true)
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleConfirmed_s);
  activity.setConditionalFormatRules(rules);

  
   // EDOI: Scheduled
  var ruleConfirmed_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Scheduled'))
    .setFontColor("#0000FF")
    .setItalic(true)
    .setRanges([rangeDateRequested])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleConfirmed_d);
  activity.setConditionalFormatRules(rules);
   
  
  // STATUS: Inspected
   var ruleInspected_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Inspected"))
    .setBold(true)
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleInspected_s);
  activity.setConditionalFormatRules(rules);
  
    // DOI: Inspected
   var ruleInspected_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Inspected"))
    .setBold(true)
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleInspected_d);
  activity.setConditionalFormatRules(rules);
  
  
    // STATUS: Comments in preparation
  var ruleComments_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Comments in preparation'))
    .setFontColor("#f00a0d")
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleComments_s);
  activity.setConditionalFormatRules(rules);
  
    // DOI: Comments in preparation
  var ruleComments_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Comments in preparation'))
    .setFontColor("#f00a0d")
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleComments_d);
  activity.setConditionalFormatRules(rules);
  
  
   // STATUS: Declined 
   var ruleDeclined_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Declined'))
    .setBackground("#B7B7B7")
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleDeclined_s);
  activity.setConditionalFormatRules(rules);
  
    // EDOI: Declined 
   var ruleDeclined_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Declined'))
    .setBackground("#B7B7B7")
    .setRanges([rangeDateRequested])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleDeclined_d);
  activity.setConditionalFormatRules(rules);
  
  
  // STATUS: Cancelled
  var ruleCancelled_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Cancelled'))
    .setBackground("#B7B7B7")
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleCancelled_s);
  activity.setConditionalFormatRules(rules);
  
  // EDOI: Cancelled
  var ruleCancelled_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus('Cancelled'))
    .setBackground("#B7B7B7")
    .setRanges([rangeDateRequested])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleCancelled_d);
  activity.setConditionalFormatRules(rules);

    
  // STATUS: Awaiting status
  var ruleAwaiting_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Awaiting status"))
    .setBold(true)
    .setFontColor("#ff9900")
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleAwaiting_s);
  activity.setConditionalFormatRules(rules);
  
    // DOI: Awaiting status
  var ruleAwaiting_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Awaiting status"))
    .setBold(true)
    .setFontColor("#ff9900")
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleAwaiting_d);
  activity.setConditionalFormatRules(rules);
  
  // STATUS: Evaluation Completed
  var ruleEvalCompleted_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Evaluation completed"))
    .setBackground("#d9ead3")
    .setFontColor("#9999a6")
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_s);
  activity.setConditionalFormatRules(rules);
  
  // DOI: Evaluation Completed
  var ruleEvalCompleted_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Evaluation Completed"))
    .setBackground("#d9ead3")
    .setFontColor("#9999a6")
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_d);
  activity.setConditionalFormatRules(rules);
  
  // Remote SIRE/OVID
  
    // STATUS: Under inspector's review
  var ruleEvalCompleted_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Under inspector's review"))
    .setFontColor("#a64d79")
    .setBold(true)
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_s);
  activity.setConditionalFormatRules(rules);
  
  // DOI: Under inspector's review
  var ruleEvalCompleted_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Under inspector's review"))
    .setFontColor("#a64d79")
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_d);
  activity.setConditionalFormatRules(rules);
  
    // STATUS: Submission in progress
  var ruleEvalCompleted_s = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Submission in progress"))
    .setFontColor("#7f6000")
    .setBold(true)
    .setRanges([rangeProcessStatus])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_s);
  activity.setConditionalFormatRules(rules);
  
  // DOI: Submission in progress
  var ruleEvalCompleted_d = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(getInspectionStatus("Submission in progress"))
    .setFontColor("#7f6000")
    .setRanges([rangeDateInspected])
    .build();
  var rules = activity.getConditionalFormatRules();
  rules.push(ruleEvalCompleted_d);
  activity.setConditionalFormatRules(rules);
   
    
};
