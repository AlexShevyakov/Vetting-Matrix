
// This function hides cases when status "closed" is selected. 
// getRange(row, column, numRows, numColumns)

// This function fires onOPen();


// HIDES ALL THE CLOSED ROWS with one click of the button
function AppealsHideAll() {
  var app = SpreadsheetApp.getActive().getSheetByName('Appeals');
  app.getRange('K3:K158').getValues().forEach(function (r, i) {
    if (r[0] !== "")
      app.hideRows(i + 1)
      });   
}

// SHOWS ALL THE CLOSED ROWS with one click of the button
function AppealsShowAll() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Appeals");
  var fullSheetRange = activeSheet.getRange(1,1,activeSheet.getMaxRows());
  activeSheet.unhideRow(fullSheetRange); 
}
  
  
  
  
  
  
  
  
  
  

