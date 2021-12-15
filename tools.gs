// Here I am planning to keep all the tools for the project


// Finding Last row in any sheet

function lastRowOfDataByCol_(arr, n){
  var i = arr.length;
  while (i--){
    if(arr[i][n])
      return i;
  }
  return 0;
};

// Get to the last row of a sheet when opening a sheet

function moveToLastRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Archive');
  let lastRow = ss.getLastRow();
  if (sheet.getMaxRows() == lastRow) {
    sheet.appendRow([""]);
  }
  lastRow = lastRow + 1;
  var range = sheet.getRange("A" + lastRow + ":A" + lastRow);
  sheet.setActiveRange(range);
  }



function traceDependents(){
  var dependents = []
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentCell = ss.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var range = ss.getDataRange();

  var regex = new RegExp("\\b" + currentCellRef + "\\b");
  var formulas = range.getFormulas();

  for (var i = 0; i < formulas.length; i++){
    var row = formulas[i];

    for (var j = 0; j < row.length; j++){
      var cellFormula = row[j].replace(/\$/g, "");
        if (regex.test(cellFormula)){
          dependents.push([i,j]);
      }
    }
  }

  var dependentRefs = [];
  for (var k = 0; k < dependents.length; k ++){
    var rowNum = dependents[k][0] + 1;
    var colNum = dependents[k][1] + 1;
    var cell = range.getCell(rowNum, colNum);
    var cellRef = cell.getA1Notation();
    dependentRefs.push(cellRef);
  }
  var output = "Dependents: ";
  if(dependentRefs.length > 0){
    output += dependentRefs.join(", ");
  } else {
    output += " None";
  }
  currentCell.setNote(output);
}