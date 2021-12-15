// This function will copy data from Activity to GreenAward sheet when GA inspection was carried out


// consider e here too. 
// function greenAwardUpdate(e) {

//if (
//    e.range.getSheet().getName() === 'Activity' &&
//    e.range.columnStart == 2 &&
//    e.range.columnEnd == 2 &&
//    e.range.rowStart >= 2 &&
//    e.range.rowEnd <= 100 &&
//    e.range.offset(0, -1).getValue() === ''
//  ) {


function greenAwardUpdate() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activity = ss.getSheetByName('Activity');
  var targetSheet = ss.getSheetByName('GreenAward');
  var greenAward = targetSheet.getDataRange().getValues();
  var firstRow = 2;
  var lastRow = 101;
  var activityData = activity.getRange(firstRow, 1, lastRow, 19).getValues();
  
  for (var i = 0; i < activityData.length; i++) {
    var vesselName = activityData[i][1];
    var survPort = activityData[i][3];
    var survType = activityData[i][4];
    var survDate = activityData[i][8];
    var survStatus = activityData[i][14];
    var survObs = activityData[i][15];
    
    // 
    if (survType == 'Green Award' && (survStatus == 'Inspected' || survStatus == 'Comments in preparation' || survStatus == 'Awaiting status' || survStatus == 'Evaluation completed')){
      
      var lastRowGreenAward = lastRowOfDataByCol_ (greenAward, 1) + 1;
      // list of vessels
      var greenVessels = targetSheet.getRange(2, 1, lastRowGreenAward, 1).getValues(); 
      var row = 0;
      
      for (var i = 0; i < greenVessels.length; i++) {
        if (greenVessels[i] == vesselName) {
          row = i + 2;
          break;
        };    
      };
      
      if (row == 0) {
        row = lastRowOfDataByCol_ (greenAward, 1) + 2;
        targetSheet.getRange(row, 1).setValue(vesselName);
        
      };
      
      targetSheet.getRange(row, 6).setValue(survDate);
      targetSheet.getRange(row, 8).setValue(survPort);
      targetSheet.getRange(row, 9).setValue(survObs);
    } 
  }
};
