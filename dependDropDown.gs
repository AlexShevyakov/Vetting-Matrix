// This is a dependable dropdown function for selecting SIRE OMs and PSC MOUs. 
// SHEET: Activity

// * This is an excellent example of how to limit onEdit range. by this line we say IF YOU EDIT WITHIN THE RANGE COL 1, ROW > 1, THEN FIRE ONEDIT.

// IMPORTANT - THIS IS AN OBSELETE SCRIPT AND IS NOT CURRENTLY USED. IT REMAINS HERE FOR LEGACY ONLY.


function dependableDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet().getName();
  var activity = ss.getSheetByName("Activity");
  var sheetData = ss.getSheetByName("$SheetData");
  var activeCell = activity.getActiveCell();
  
  if(activeCell.getColumn() == 5 && sheet == "Activity" && activeCell.getRow() > 1){
//    activeCell.offset(0, 2).clearContent().clearDataValidations();   
    
    var inspectionTypes = sheetData.getRange(1, 8, 1, 8).getValues(); // ANY CHANGES TO the "$SheetData" WILL HAVE TO BE REFLECTED HERE!!!
    
  
    
    var selectedValue = activeCell.getValue();
    var inspectionTypesIndex = inspectionTypes[0].indexOf(selectedValue) + 1;
    
    var validationRange = sheetData.getRange(2, inspectionTypesIndex + 7, sheetData.getLastRow()); // inspectionTypesIndex + 7 
    
//    Logger.log(inspectionTypesIndex);
    
    var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
    activeCell.offset(0, 2).clearContent().clearDataValidations().setDataValidation(validationRule);
  };
};





