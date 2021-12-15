
// This function hides cases when status "closed" is selected. 
// getRange(row, column, numRows, numColumns)

// This function fires onOPen();


// HIDES ALL THE CLOSED ROWS with one click of the button
function screeningHideAll() {
  var app = SpreadsheetApp.getActive().getSheetByName('Screening');
  app.getRange('G:G').getValues().forEach(function (r, i) {
    if (r[0] == "Closed")
      app.hideRows(i + 1)
      });   
}

// SHOWS ALL THE CLOSED ROWS with one click of the button
function screeningShowAll() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Screening");
  var fullSheetRange = activeSheet.getRange(1,1,activeSheet.getMaxRows());
  activeSheet.unhideRow(fullSheetRange); 
}
  
  
  
  
  
  
  
  
  
  

