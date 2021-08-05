function printGrafo() {
  var html = HtmlService.createHtmlOutputFromFile('index').setHeight(664.3).setWidth(1294.3);
  SpreadsheetApp.getUi().showModelessDialog(html, "Mapa Conceitual Estendido");
}