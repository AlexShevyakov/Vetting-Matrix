function onOpen() {
  menu();
  alertRequestOverdue();
  alertCommentsDue();
 // alertPendingReminder();
  clearComments();
  normalizeRowsCount();

};


 

function menu (){
    var ui = SpreadsheetApp.getUi();
  ui.createMenu('Vetting Team')
  .addItem('Sort ACTIVITY by DOI', 'sortByDOI')
  .addItem('Sort ACTIVITY by Vessels Name', 'SortByVesselName')
  .addItem('Rejections - Hide all "closed"', 'RejectionsHideAll')
  .addItem('Rejections - Show all', 'RejectionsShowAll')
      .addSeparator()
  
  .addItem('Close Process', 'closeProcess')
      .addSeparator()
  
  .addItem('Update Fleet Status', 'fleetStatusUpdate')
  .addItem('Update missing IDs', 'printMissingID')
      .addSeparator()

  .addItem('Sort ARCHIVE by DOI', 'sortByDOI_Archive')
  .addItem('Sort ARCHIVE by Vessels Name', 'sortByVessel_Archive')
  .addItem('Check ARCHIVE for duplicated records', 'colouredDuplicates')
      .addSeparator()

  .addItem('Sort INCIDENTS by Date of insidents', 'sortByDOI_Incidents')
  .addItem('Sort INCIDENTS by Vessels Name', 'sortByVessel_Incidents')
  .addItem('Hide closed INCIDENTS', 'IncidentsHideAll')
  .addItem('Show all Incidents', 'IncidentsShowAll')
      .addSeparator()

  .addItem('Hide closed APPEALS', 'AppealsHideAll')
  .addItem('Show all APPEALS', 'AppealsShowAll')
  .addItem('Sort APPEALS by Date of inspection', 'sortByDOI_appeals')
  .addItem('Sort APPEALS by Vessels Name', 'sortByVessel_appeals')
      .addSeparator()

  .addItem('Hide closed SCREENINGS', 'screeningHideAll')
  .addItem('Show all SCREENINGS', 'screeningShowAll')
      .addSeparator()

  .addItem('Run Acceptance verifications', 'acceptanceStatusColour')
        .addSeparator()
        
  .addItem('Sort COVID19', 'sortByDateModified_covid19')
  .addItem('Hide COVID19 rows behind 2 weeks old', 'hideCOVIDrows')
    .addSeparator()

  .addItem('Trace dependencies','traceDependents')
    .addSeparator()
  
  .addItem('Export Charts to PNG', 'exportCharts')
      .addToUi()
 
};

