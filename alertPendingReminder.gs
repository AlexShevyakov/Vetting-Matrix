/* This function colours the value in column "Comments submitted" if Awaiting Status has been idle for 2 weeks.
// Conditions for the alert to activate are:

*/

// This function runs onOpen()
// SHEET: ACTIVITY
// Module version: 03-Apr-2019


/*
function alertPendingReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var activity = ss.getSheetByName('Activity');
  var submittedColumn = activity.getRange(2, 12, activity.getLastRow()-1, 1);
  var submitDate = submittedColumn.getValues();
  var statusColumn = activity.getRange(2, 15, activity.getLastRow()-1, 1);
  var status = statusColumn.getValues();
  var companyColumn = activity.getRange(2, 7, activity.getLastRow()-1, 1);
  var company = companyColumn.getValues();
  
  activity.getRange("L2:L100").setBorder(false, false, false, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID).clearNote();
  
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24; // Exactly 1 day, 24 hours
  var today = new Date();
  
  for (var i = 0; i < submitDate.length; i++) {
    var remindDate = new Date(submitDate[i][0]);
    var reminder_1 = new Date(remindDate.getTime() + 14 * MILLIS_PER_DAY); // for common inspections
    var reminder_2 = new Date(remindDate.getTime() + 30 * MILLIS_PER_DAY); // for LUKOIL and POT inspections
    
    if (company[i][0] !== "LUKOIL" && company[i][0] !== "PRIMORSK OIL"){
      if (!remindDate.isBlank && reminder_1 < today && status[i][0] == "Awaiting status"){
        activity.getRange(i + 2, 12, 1, 1).setBorder(true, true, true, true, false, false, '#f7ac20 ', SpreadsheetApp.BorderStyle.SOLID_THICK).setNote("Status is being awaited longer than 14 days"); // ORANGE
      } 
    } else if 
      (!remindDate.isBlank && reminder_2 < today && status[i][0] == "Awaiting status"){
        activity.getRange(i + 2, 12, 1, 1).setBorder(true, true, true, true, false, false, '#ff0000', SpreadsheetApp.BorderStyle.SOLID_THICK).setNote("Status is being awaited longer than 30 days"); // RED
      };
  };
};


*/

