function printGrafo() {
  // Criei as duas formas indicadas na documentação 
  var html = HtmlService.createHtmlOutputFromFile('index').setHeight(920).setWidth(1900);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModelessDialog(html, 'MCE');
}