// Date last modified: 28/3/2019

// Sorting Activity by the Date of Inspection
function sortByDOI() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Activity");
  activeSheet.getRange("A2:U100").activate();
  activeSheet.sort(9, true); // Column two = by Date of Inspection
};

// SORTING OF ACTIVITY PAGE

function SortByVesselName() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Activity");
  activeSheet.getRange("A2:U100").activate();
  activeSheet.sort(2, true); // Column 2(B) = by Vessel's name
  alertPendingReminder();
};

// SORTING OF ARCHIVE by DOI

function sortByDOI_Archive() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Archive");
  activeSheet.getRange("A2:K").activate();
  activeSheet.sort(8, true); // Column 8(H) = by Date of Inspection
};

// SORTING OF ARCHIVE by Vessel Name
function sortByVessel_Archive() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Archive");
  activeSheet.getRange("A2:K").activate();
  activeSheet.sort(2, true); // Column 2(B) = by Vessel Name
};

// SORTING COVID19 by date modified
/// NEW APPROACH

function sortByDateModified_covid19() {
  
  //Variable for column to sort first
  
  var sortFirst = 7; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = false; //Set to false to sort descending
  
  //Variables for column to sort second
 
 // var sortSecond = 3;
 // var sortSecondAsc = false;
  
  //Number of header rows
  
  var headerRows = 1; 

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('COVID19'); //name of sheet to be sorted
  var range = sheet.getRange(headerRows + 1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());
  range.sort([{column: sortFirst, ascending: sortFirstAsc}]);
  hideCOVIDrows();
}


// SORTING OF INCIDENTS by date

function sortByDOI_Incidents() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Incidents");
  activeSheet.getRange("A3:O").activate();
  activeSheet.sort(4, true); // Column 4(D) = by Date of Incident
};

// SORTING OF INCIDENTS by Vessel Name
function sortByVessel_Incidents() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getSheetByName("Incidents");
  activeSheet.getRange("A3:O").activate();
  activeSheet.sort(1, true); // Column 1(A) = by Vessel Name
};

// SORTING APPEALS by date of inspection
/// NEW APPROACH

function sortByDOI_appeals() {
  
  //Variable for column to sort first
  
  var sortFirst = 4; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = false; //Set to false to sort descending
  var headerRows = 1; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appeals');
  var range = sheet.getRange(headerRows + 1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());
  range.sort([{column: sortFirst, ascending: sortFirstAsc}]);
}

// SORTING APPEALS by date modified
/// NEW APPROACH

function sortByVessel_appeals() {
  
  //Variable for column to sort first
  
  var sortFirst = 1; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = true; //Set to false to sort descending
  var headerRows = 3; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Appeals');
  var range = sheet.getRange(headerRows + 1, 1, sheet.getMaxRows()-headerRows, 2);
  range.sort([{column: sortFirst, ascending: sortFirstAsc}]);
}

// Sorts Due Tab by the Due range THEN by the vessel name
function sortDue() {
  const app = SpreadsheetApp;
  const ss = app.getActiveSpreadsheet().getSheetByName("DueData");
  ss.getRange('A3:C60').activate()
    .sort([{ column: 3, ascending: false }, { column: 1, ascending: true }]);
};


