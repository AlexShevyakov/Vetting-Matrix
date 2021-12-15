// Non Array columns: [A-1]	[B-2]	[C-3]	[D-4]	[E-5]	[F-6]	[G-7]	[H-8]	[I-9]	[J-10]	[K-11]	[L-12]	[M-13]	[N-14]	[O-15]	[P-16]	[Q-17]	[R-18]	[S-19]	[T-20]	[U-21]	[V-22]	[W-23]	[X-24]	[Y-25]	[Z-26]
// Array columns: [A-0]	[B-1]	[C-2]	[D-3]	[E-4]	[F-5]	[G-6]	[H-7]	[I-8]	[J-9]	[K-10]	[L-11]	[M-12]	[N-13]	[O-14]	[P-15]	[Q-16]	[R-17]	[S-18]	[T-19]	[U-201]	[V-21]	[W-22]	[X-23]	[Y-24]	[Z-25]
// Array row: 1-0

/*

THIS MODULE DEPRECIATED DUE TO CHANGE OF WORKFLOW - WE USE FORMULAS NOW AND NOT SCRIPT TO UPDATE FLEET STATUS
SHEET: FleetStatus
This function will colour cells depending on the Acceptance of an inspection - green for Acceptable and Red for Unacceptable
*/

function acceptanceStatusColour() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activity = ss.getSheetByName("Activity");
  var activityRange = activity.getRange(1, 1, activity.getLastRow(), 20).getValues();
  var fleetStatusRange

  // Added to ensure we are working on the specific range, and static range
  var activeCell = activity.getActiveCell();
  var sheet = ss.getActiveSheet().getName();


  var vessel = activity.getRange(1, 2, activity.getLastRow(), 1).getValues();
  var company = activity.getRange(1, 7, activity.getLastRow(), 1).getValues();

  // Condition to ensure we are editing only OUTCOME on the ACTIVITY sheet
  var activeCol = activeCell.getColumn();


  if (sheet == "Activity" && activeCol == 20) {
    for (var i = 0; i < activityRange.length; i++) {
      var inspStatus = activityRange[i][16];
      var inspType = activityRange[i][4];
      var vesselName = activityRange[i][1];
      var acceptanceOutcome = activityRange[i][20];

      // SIRE and CDI inspections
      // If Acceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Acceptable" && (inspType == "SIRE" || inspType == "Remote SIRE" || inspType == "CDI")) {
        var column = findCompanySIRE(company[i]);
        var row = findVesselSIRE(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_sire").getRange(row, column).setBackground("#6AA84F");
        };
      };
      // If Unacceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Unacceptable" && (inspType == "SIRE" || inspType == "Remote SIRE" || inspType == "CDI")) {
        var column = findCompanySIRE(company[i]);
        var row = findVesselSIRE(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_sire").getRange(row, column).setBackground("#E06666");
        };
      };



      // OVID group
      // If Acceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Acceptable" && inspType == "OVID" || inspType == "Remote OVID") {
        var column = findCompanyOVID(company[i]);
        var row = findVesselOVID(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_ovid").getRange(row, column).setBackground("#6AA84F");
        };
      };
      // If Unacceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Unacceptable" && inspType == "OVID") {
        var column = findCompanyOVID(company[i]);
        var row = findVesselOVID(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_ovid").getRange(row, column).setBackground("#E06666");
        };
      };



      // DRY-BULK group
      // If Acceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Acceptable" && (vesselName == "ESL AMERICA" || vesselName == "NORDIC OASIS" || vesselName == "NORDIC ODIN" || vesselName == "NORDIC ODYSSEY" || vesselName == "NORDIC OLYMPIC" || vesselName == "NORDIC ORION" || vesselName == "NORDIC OSHIMA" || vesselName == "NS ENERGY" || vesselName == "NS YAKUTIA")) {
        var column = findCompanyDRY(company[i]);
        var row = findVesselDRY(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_drybulk").getRange(row, column).setBackground("#6AA84F");
        };
      };
      // If Unacceptable
      if (inspStatus == "Evaluation completed" && acceptanceOutcome == "Unacceptable" && (vesselName == "ESL AMERICA" || vesselName == "NORDIC OASIS" || vesselName == "NORDIC ODIN" || vesselName == "NORDIC ODYSSEY" || vesselName == "NORDIC OLYMPIC" || vesselName == "NORDIC ORION" || vesselName == "NORDIC OSHIMA" || vesselName == "NS ENERGY" || vesselName == "NS YAKUTIA")) {
        var column = findCompanyDRY(company[i]);
        var row = findVesselDRY(vessel[i]);
        if (column > 0 && row > 0) {
          ss.getSheetByName("FS_drybulk").getRange(row, column).setBackground("#E06666");
        };
      }
    };
  };
};


// SIRE and CDI group
// Finds index of the Company column
function findCompanySIRE(company) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_sire");
  var lastColumn = fleetStatus.getLastColumn();
  var cmp = company.toString();

  var data = fleetStatus.getRange(1, 1, 1, lastColumn).getValues();
  data = data[0];//Get the first and only inner array

  var companyIndex = data.indexOf(cmp) + 1;//Arrays are zero indexed- add 1

  return companyIndex;
};

// Finds the row with the vessel's name
function findVesselSIRE(vessel) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_sire");
  var lastRow = fleetStatus.getLastRow();
  var vsl = vessel.toString();

  var data = fleetStatus.getRange(1, 1, lastRow, 1).getValues();//Get 2D array of all values in row one

  for (var i = 0; i < data.length; i++) {
    if (data[i] == vsl) {
      return i + 1;
    }
  }
  return 0;
};

// OVID group
// Finds index of the Company column
function findCompanyOVID(company) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_ovid");
  var lastColumn = fleetStatus.getLastColumn();
  var cmp = company.toString();

  var data = fleetStatus.getRange(1, 1, 1, lastColumn).getValues();
  data = data[0];//Get the first and only inner array

  var companyIndex = data.indexOf(cmp) + 1;//Arrays are zero indexed- add 1

  return companyIndex;
};

// Finds the row with the vessel's name
function findVesselOVID(vessel) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_ovid");
  var lastRow = fleetStatus.getLastRow();
  var vsl = vessel.toString();

  var data = fleetStatus.getRange(1, 1, lastRow, 1).getValues();//Get 2D array of all values in row one

  for (var i = 0; i < data.length; i++) {
    if (data[i] == vsl) {
      return i + 1;
    }
  }
  return 0;
};

// DRY BULK group
// Finds index of the Company column
function findCompanyDRY(company) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_drybulk");
  var lastColumn = fleetStatus.getLastColumn();
  var cmp = company.toString();

  var data = fleetStatus.getRange(1, 1, 1, lastColumn).getValues();
  data = data[0];//Get the first and only inner array

  var companyIndex = data.indexOf(cmp) + 1;//Arrays are zero indexed- add 1

  return companyIndex;
};

// Finds the row with the vessel's name
function findVesselDRY(vessel) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fleetStatus = ss.getSheetByName("FS_drybulk");
  var lastRow = fleetStatus.getLastRow();
  var vsl = vessel.toString();

  var data = fleetStatus.getRange(1, 1, lastRow, 1).getValues();//Get 2D array of all values in row one

  for (var i = 0; i < data.length; i++) {
    if (data[i] == vsl) {
      return i + 1;
    }
  }
  return 0;
};
