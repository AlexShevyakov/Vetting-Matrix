
// This function hides cases when status "closed" is selected. 
// getRange(row, column, numRows, numColumns)

// This function fires onOPen();

function hideCOVIDrows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const covid19 = ss.getSheetByName('COVID19');
  const lastRow = covid19.getLastRow();
  const dates = covid19.getRange(2, 7, lastRow, 1).getValues(); 
  
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24; // Exactly 1 day or 24 hours
  const today = new Date();
  const barrier = new Date(today - 14 * MILLIS_PER_DAY);

  for (let i = 0; i < dates.length; i++) { 
    if (dates[i][0] !=="" && dates[i][0] < barrier) { 
      covid19.hideRows(2 + i);
    } 
  } 
}
  
  
  
  
  
  
  
  
  
  

