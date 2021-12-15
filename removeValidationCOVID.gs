// This function will remove validation and convert values to statis ones in sheet COVID19.
// getRange(row, column, numRows, numColumns) 


// Array has first position at 0
// Range has first position at 1

function removeValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const covid19 = ss.getSheetByName("COVID19");
  
  let lastRow = covid19.getLastRow();
  let lastCol = 7;
  const dataRange = covid19.getRange(2, 1, lastRow, lastCol);
  const data = dataRange.getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][5] !== "") {
      covid19.getRange(i + 2, 1).clearDataValidations();
      let columnsToText = covid19.getRange(i + 2, 2, 1, 4);
      columnsToText.copyTo(columnsToText, {contentsOnly: true});           
    }
  }
}

// Automatically adds timestamp to the column "Modified Date"

function timeStampCOVID(e) {
  var ss = e.source.getActiveSheet();
  var sheet = ['COVID19'];
  var row = e.range.getRow();
  
  // Columns with the data to be tracked; 1-based
  var ind = [6].indexOf(e.range.columnStart); 
  
  // The columb being updated, 1-based
  var stampCols = [7]
  
  if(sheet.indexOf(ss.getName()) == -1 || ind == -1) 
    return; 
  
  if (e.source.getSheetByName('COVID19').getRange(row, 7).getValue() == '') { // Checking if there is a value in the column
    
    // Insert/Update the timestamp.
    var timestampCell = ss.getRange(e.range.rowStart, stampCols[ind]);
    timestampCell.setValue(typeof e.value == 'object' ? null :  new Date());
    removeValidation();
    sortByDateModified_covid19();
  };
};