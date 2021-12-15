
// This function hides cases when status "closed" is selected. 
// This function fires onEdit();
//SHEET: RejectHold
function rejectClose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var sheet = "RejectHold";
  
  //Ensure on correct cheet.
  if(sheet == activeSheet.getName()){
    var cell = ss.getActiveCell()
    var cellValue = cell.getValue();
    
    //Ensure we are looking at the correct column.
    if(cell.getColumn() == 12){
      //If the cell matched the value we require, hide the row. 
      if(cellValue == "Closed"){
        activeSheet.hideRow(cell);
        formatClosed();
      };
    };
  };
}


// SHEET: REJECTIONS.
// HIDES ALL THE CLOSED ROWS with one click of the button
function RejectionsHideAll() {
  var app = SpreadsheetApp.getActive().getSheetByName('RejectHold');
  app.getRange('L:L').getValues().forEach(function (r, i) {
    if (r[0] == "Closed")
      app.hideRows(i + 1)
      });   
}

// SHEET: REJECTIONS.
// SHOWS ALL THE CLOSED ROWS with one click of the button
function RejectionsShowAll() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("RejectHold");
  var fullSheetRange = activeSheet.getRange(1,1,activeSheet.getMaxRows());
  activeSheet.unhideRow(fullSheetRange); 
}