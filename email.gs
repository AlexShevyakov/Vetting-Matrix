
/**
* Sends emails with data from the current spreadsheet.
*/
function sendEmails() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("$SheetData");
  
  var startRow = 2; // First row of data to process
  var numRows = 2; // Number of rows to process
  // Fetch the range of cells A2:B3
  var dataRange = sheet.getRange(startRow, 17, numRows, 2);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i in data) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = row[1]; // Second column
    var subject = 'Sending emails from a Spreadsheet';
    MailApp.sendEmail(emailAddress, subject, message);
  }
}

// You need this solution 
// https://stackoverflow.com/questions/36529890/emailing-google-sheet-range-with-or-without-formatting-as-a-html-table-in-a-gm

