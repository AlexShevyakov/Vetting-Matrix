// This is Global onEdit script for the entire project.
// This is the only way to maintain multiple onEdit()'s within a signle project 

function onEdit(e){
  generateID(e);
  unique_ID_screenings(e);
  unique_ID_inciednts(e);
  timeStampCOVID(e);
//  greenAwardUpdate(); Think how to localize the firing up!
 // sortByDOI_Archive();
 // acceptanceStatusColour();
  removeValidationAppeals();
  rejectClose();
  incidentClose();

};