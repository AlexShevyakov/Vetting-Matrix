
function caseID(len) {
  let u_str = ''
  while (u_str.length < len) u_str += Math.random().toString(36).substr(2, len - u_str.length);
  return u_str;
}


// THis is Legacy from the time when you did not know that multiple onEdit() are not allowed.
// The function has been moved into GLOBAL onEdit() 

function generateID(e) {
  // This block limits the use of the onEdit function - it will only fire if changes are made within the range specified.
  if (
    e.range.getSheet().getName() === 'Activity' &&
    e.range.columnStart == 2 &&
    e.range.columnEnd == 2 &&
    e.range.rowStart >= 2 &&
    e.range.rowEnd <= 100 &&
    e.range.offset(0, -1).getValue() === ''
  ) {
    e.value !== '' ? e.range.offset(0, -1).setValue(caseID(8)) : null;
  }
}

// this function print IDs into a blank sheet 'IDs'
function printIDs() {
  const app = SpreadsheetApp;
  const printSheet = app.getActiveSpreadsheet().getSheetByName("IDs");
//  const range = printSheet.getRange(1, 1, 100, 1).getValues();
  
    for (let i = 0; i < 100; i++) {
      printSheet.getRange(i + 1, 1).setValue(caseID(8))
     };
};

// this function checks ACTIVITY the first two columns and if ID is missing - inserts it.

function printMissingID() {
  const app = SpreadsheetApp;
  const printSheet = app.getActiveSpreadsheet().getSheetByName("Activity");
  const lastRow = printSheet.getLastRow();
  const checkRange = printSheet.getRange(2, 1, lastRow, 2).getValues();
  
    for (let i = 0; i < checkRange.length; i++) {
      if(checkRange[i][0] == "" && checkRange[i][1] !== "") {
      printSheet.getRange(i + 2, 1).setValue(caseID(8))
     };
};
};
