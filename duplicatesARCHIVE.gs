// This mega code finds duplicate rows and mark them in Archive

/**
 * The entry point.
 */
function colouredDuplicates(){
  
  // sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName('Archive');
  var collection = sheet.getDataRange().getValues().slice(1);
  // validate data
  var hasUndefined = validateSheetData_(collection);
  if(hasUndefined !== -1) {
    SpreadsheetApp.getActive().toast('No such vessel in the FleetData - check spelling. Index: ' + hasUndefined + collection[hasUndefined][1], '', -1);
    sheet.getRange(hasUndefined + 2, 2).activate();
    return;
  }
  // finds dublicates
  var rows = duplicatesRow(collection);
  paintDublicatesRowSheet(sheet, rows);
};

// JSON.stringify(collection, null, ' ')

function validateSheetData_(collection){
   var fleetDataIndex = getfleetDataIndex();
   return collection.findIndex(function(row){
    return row[1] !== '' && !existInFleetData(row[1], fleetDataIndex);
  });
}


/**
* @returns {object[]}
*/
function duplicatesRow(collection) {
  // $UNRESOLVED - need to have rowNumber shows when inconsistency found
  var shiftSize = 2;
  collection = collection.reduce(function(p, v, i){
    var index = v[1] + v[4] + (v[7].getTime ? v[7].getTime() : '_'); // Checking Col B, E and H - they should be the same to cause warning.
    if(!p.hasOwnProperty(index)){
      p[index] = {
        rows: []
      };
    }
    p[index].rows.push(i + shiftSize);
    return p;
  }, {});
  var res = [];
  for(var k in collection){
    if(collection.hasOwnProperty(k))
      if(collection[k].rows.length > 1)
        res = res.concat(collection[k].rows.filter(function(i){ return res.indexOf(i) === -1 }));
  }
  
  return res;
};

function paintDublicatesRowSheet(sheet, rows){
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var colors = [];
  for(var i = 1; i <= lastRow; i++){
    if(rows.indexOf(i) === -1){
      colors.push(prefillArray(lastColumn, ''));
    } else {
      colors.push(prefillArray(lastColumn, '#EAD1DC'));
    }
  }
  return sheet.getRange(1, 1, lastRow, lastColumn).setBackgrounds(colors);
};

function prefillArray(length, value){
  return Array.apply(null, Array(length)).map(String.prototype.valueOf,value);
};

// duplicatesRow debug

function existInFleetData(vessel, fleetDataIndex){
  return fleetDataIndex.indexOf(vessel) !== -1
}

function getfleetDataIndex(){
  var sheet = SpreadsheetApp.getActive().getSheetByName('$FleetData');
  return sheet.getDataRange().getValues().slice(sheet.getFrozenRows() || 1).map(function(row){
    return row[0];
  });
}

function duplicatesRowDebug () {
  duplicatesRow(SpreadsheetApp.getActiveSheet());
}