// Non Array columns: [A-1]	[B-2]	[C-3]	[D-4]	[E-5]	[F-6]	[G-7]	[H-8]	[I-9]	[J-10]	[K-11]	[L-12]	[M-13]	[N-14]	[O-15]	[P-16]	[Q-17]	[R-18]	[S-19]	[T-20]	[U-21]	[V-22]	[W-23]	[X-24]	[Y-25]	[Z-26]
// Array columns: [A-0]	[B-1]	[C-2]	[D-3]	[E-4]	[F-5]	[G-6]	[H-7]	[I-8]	[J-9]	[K-10]	[L-11]	[M-12]	[N-13]	[O-14]	[P-15]	[Q-16]	[R-17]	[S-18]	[T-19]	[U-201]	[V-21]	[W-22]	[X-23]	[Y-24]	[Z-25]
// Array row: 1-0
// getRange(row, column, numRows, numColumns) 
// getRange(row, column) 

// SHEET: Activity
// This functoin will move completed inspections to the Archive and clears the Activity row.

function closeProcess() {
  
    
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activity = ss.getSheetByName("Activity");
  
  normalizeRowsCount(activity);
   
  var activityRange = activity.getRange(1, 1, activity.getLastRow(), activity.getLastColumn()).getValues();
  
  var counterTransfer = 0;
  var counterDelete = 0;
  
  for (var i = activityRange.length - 1; i > 0; i--) {   
    if (activityRange[i][18] == true && activityRange[i][15] == "Evaluation completed") {
   
      var targetSheet = ss.getSheetByName("Archive");
     
      // Three sets because we copy selectied columns only.
       
      var activitySet1 = activity.getRange(i + 1, 1, 1, 2); //   Case ID, Vessel
      var activitySet2 = activity.getRange(i + 1, 4, 1, 4); //   Port, Type of inspection, Operation, Inspected by
      var activitySet3 = activity.getRange(i + 1, 9, 1, 1); //   Date of inspection
      var activitySet4 = activity.getRange(i + 1, 17, 1, 1); //  Observations
      var activitySet5 = activity.getRange(i + 1, 20, 1, 1); //  Outcome
      
      var targetRangeSet1 = targetSheet.getRange(lastRowArchive() + 1, 1);
      var targetRangeSet2 = targetSheet.getRange(lastRowArchive() + 1, 4);
      var targetRangeSet3 = targetSheet.getRange(lastRowArchive() + 1, 8);
      var targetRangeSet4 = targetSheet.getRange(lastRowArchive() + 1, 9);
      var targetRangeSet5 = targetSheet.getRange(lastRowArchive() + 1, 10);
            
      activitySet1.copyTo(targetRangeSet1,{contentsOnly:true});
      activitySet2.copyTo(targetRangeSet2,{contentsOnly:true});
      activitySet3.copyTo(targetRangeSet3,{contentsOnly:true});
      activitySet4.copyTo(targetRangeSet4,{contentsOnly:true});
      activitySet5.copyTo(targetRangeSet5,{contentsOnly:true});
      
      // activity.deleteRow(i + 1);
      activity.getRange(i + 1, 1, 1, 20).clearContent();
      counterTransfer ++; // Update counter
  };
     if(activityRange[i][18] == true && activityRange[i][15] == "Cancelled" || activityRange[i][15] == "Declined") {
      // activity.deleteRow(i + 1);
      activity.getRange(i + 1, 1, 1, 20).clearContent();
      counterDelete ++; // Update counter
     }; 
  }; 
  if (counterTransfer > 0){
    completeMsgTransfer(counterTransfer);
  };
  if(counterDelete > 0){
    completeMsgDelete(counterDelete)
  };
};

/**
 * The function description
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Activity sheet for normalization
 */
function normalizeRowsCount(activity){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activity = ss.getSheetByName("Activity");
  var maxRow = activity.getMaxRows();
  if (maxRow < 101)
   activity.insertRowsBefore(maxRow - 1, 101 - maxRow);
}

// Functoin that finds last row in the ARCHIVE. This function is then passed to the main function "closeProcess()"
  
function lastRowArchive(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archive = ss.getSheetByName("Archive");
  var values = archive.getRange("B1:B").getValues();
  var lastValue = values.filter(String).length;
 
  return lastValue;
};


function completeMsgTransfer(counterTransfer){
   SpreadsheetApp.getUi().alert(counterTransfer + " Closed inspection(s) were copied to ARCHIVE. Process completed");
}

function completeMsgDelete(counterDelete){
   SpreadsheetApp.getUi().alert(counterDelete + " Cancelled or Declined inspection(s) were deleted from ACTIVITY");
}