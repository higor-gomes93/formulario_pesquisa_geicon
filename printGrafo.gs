function printGrafo() {
  var html = HtmlService.createHtmlOutputFromFile('index').setHeight(664.3).setWidth(1294.3);  // Tamanho do plot
  var nome = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 1).getValue();
  SpreadsheetApp.getUi().showModelessDialog(html, nome);
}