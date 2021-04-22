function insertImageOnSpreadsheet() {
  var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1-ha8jO5WTWTvBddnl47OxmDf_ou1iFMEbUIK86ijJyQ/edit#gid=2139807765';
  // Name of the specific sheet in the spreadsheet.
  var SHEET_NAME = 'Teste';

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  var response = UrlFetchApp.fetch(
      'https://docs.google.com/uc?id=1HSsOacDPENmuwpuJOJeVixTZxY_kB-YH');
  var binaryData = response.getContent();

  // Insert the image in cell A1.
  var blob = Utilities.newBlob(binaryData, 'image/png', 'MyImageName');
  sheet.insertImage(blob, 1, 1);
}