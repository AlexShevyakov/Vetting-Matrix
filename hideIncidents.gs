// This function hides INCIDENT cases when status "closed" is selected. 
// This function fires onEdit();
//SHEET: Incidents
function incidentClose() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var sheet = "Incidents";
  
  //Ensure on correct cheet.
  if(sheet == activeSheet.getName()){
    var cell = ss.getActiveCell()
    var cellValue = cell.getValue();
    
    //Ensure we are looking at the correct column.
    if(cell.getColumn() == 14){
      //If the cell matched the value we require, hide the row. 
      if(cellValue == "Closed"){
        activeSheet.hideRow(cell);
      };
    };
  };
}


// SHEET: REJECTIONS.
// HIDES ALL THE CLOSED ROWS with one click of the button
function IncidentsHideAll() {
  var app = SpreadsheetApp.getActive().getSheetByName('Incidents');
  app.getRange('O:O').getValues().forEach(function (r, i) {
    if (r[0] == "Closed")
      app.hideRows(i + 1)
      });   
}

// SHEET: REJECTIONS.
// SHOWS ALL THE CLOSED ROWS with one click of the button
function IncidentsShowAll() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Incidents");
  var fullSheetRange = activeSheet.getRange(1,1,activeSheet.getMaxRows());
  activeSheet.unhideRow(fullSheetRange); 
}