/* This function sends ALERT if there is no activity 3 days after insepction has been requested.
// Conditions for the alert to activate are:

*/

// This function runs onOpen()
// SHEET: ACTIVITY

function alertRequestOverdue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activity');
  var requestColumn = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1);
  var requestDate = requestColumn.getValues();
  var statusColumn = sheet.getRange(2, 15, sheet.getLastRow() - 1, 1);
  var status = statusColumn.getValues();

  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var today = new Date();
  var criteria = new Date(today.getTime() - 3 * MILLIS_PER_DAY);

  for (var i = 0; i < requestDate.length; i++) {
    var aDate = new Date(requestDate[i][0]);

    if (!aDate.isBlank && aDate.getTime() < criteria && status[i][0] == "Requested") {
      sheet.getRange(i + 2, 8, 1, 1).setBackground('#F46525');
    } else {
      sheet.getRange(i + 2, 8, 1, 1).setBackgroundRGB(255, 242, 204);
    }
  }
}


