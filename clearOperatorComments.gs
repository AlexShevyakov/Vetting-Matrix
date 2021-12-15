/**
 * This module clears Operator Comments when the status of inspection changes from Comments in preparation to "Awaiting status" or "Evaluation completed"
 *
 */

function clearComments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activity = ss.getSheetByName("Activity");

  const status = activity.getRange("P2:P100").getValues();


  for (let i = 0; i < status.length; i++) {

    if (status[i][0] == "Awaiting status" || status[i][0] == "Evaluation completed") {
      console.log(status[i][2])
      activity.getRange(i + 2, 18, 1, 1).setValue("");
    }
  }
}




