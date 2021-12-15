/* This function sends ALERT if comments are due to be published within 3 days.

*/
// test clasp
// This function runs onOpen()
// SHEET: ACTIVITY

function alertCommentsDue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var activity = ss.getSheetByName('Activity');
  var publishColumn = activity.getRange(2, 11, activity.getLastRow()-1, 1);
  var publishDue = publishColumn.getValues();
  var statusColumn = activity.getRange(2, 15, activity.getLastRow()-1, 1);
  var status = statusColumn.getValues();
  var inspTypeColum = activity.getRange(2, 5, activity.getLastRow()-1, 1);
  var inspType = inspTypeColum.getValues();
  
  
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24; // Exactly 1 day or 24 hours
  var today = new Date();
  activity.getRange("K2:K100").setBackground("#fff2cc").setFontColor('black');
  
  for (var i = 0; i < publishDue.length; i++) {
    var dueDate = new Date(publishDue[i][0]);
    var barrier = new Date(dueDate.getTime() - 3 * MILLIS_PER_DAY);
       
    if (!dueDate.isBlank && today > barrier && today < dueDate && inspType[i][0] == "SIRE" && (status[i][0] == "Inspected" || status[i][0] == "Comments in preparation")) {
      activity.getRange(i + 2, 11, 1, 1).setBackground("#ffff00");
      // here I need email to be sent with a table
    } else if (!dueDate.isBlank && today >= dueDate && inspType[i][0] == "SIRE" && (status[i][0] == "Inspected" || status[i][0] == "Comments in preparation")) { 
      activity.getRange(i + 2, 11, 1, 1).setBackground("#FF0000").setFontColor('white'); //Red background, white font
    } else {
      activity.getRange(i + 2, 11, 1, 1).setBackground("#fff2cc").setFontColor('black'); // Default colours
    };
  };
};


