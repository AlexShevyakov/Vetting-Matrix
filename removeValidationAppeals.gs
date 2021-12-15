// This function will remove validation and convert values to statis ones in sheet COVID19.
// getRange(row, column, numRows, numColumns) 


// Array has first position at 0
// Range has first position at 1

function removeValidationAppeals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appeals = ss.getSheetByName("Appeals");
  
  let lastRow = appeals.getLastRow();
  let lastCol = 12;
  const dataRange = appeals.getRange(2, 1, lastRow, lastCol);
  const data = dataRange.getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][5] !== "") {
      appeals.getRange(i + 2, 1).clearDataValidations();
      let columnsToText = appeals.getRange(i + 2, 2, 1, 5);
      columnsToText.copyTo(columnsToText, {contentsOnly: true});           
    }
  }
}

