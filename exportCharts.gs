
function exportCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  const charts = sheet.getCharts();
  let url;

  for (let i = 0; i < charts.length; i++) {
    // creating a proxy slide
    const proxySlide = SlidesApp.create("proxySlide " + i);
    const proxySaveSlide = proxySlide.getSlides()[0];
    const chartImage = proxySaveSlide.insertSheetsChartAsImage(charts[i]);

    // Getting image from slides
    const myimage = chartImage.getAs('image/png').setName("chart " + i);
    url = DriveApp.getFolderById("1bQ4K034w-Q57LrpxZxe3ZZZWbvIe_I8J").createFile(myimage).getUrl();
    DriveApp.getFileById(proxySlide.getId()).setTrashed(true);
  }  
  return url;
}


