function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function pegador(){
  var elements = [];
  var conceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 2, 10, 1).getValues();
  var conceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 3, 10, 1).getValues();
  var posXConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 5, 10, 1).getValues();
  var posYConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 4, 10, 1).getValues();
  var posXConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 7, 10, 1).getValues();
  var posYConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 6, 10, 1).getValues();
  var valorSinal = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual").getRange(1, 8, 10, 1).getValues();

  for(let i = 0; i < 10; i++){
    let nodeCC = { data: { id: conceitoCausa[i][0] }, position: { x: parseInt(posXConceitoCausa[i][0]), y: parseInt(posYConceitoCausa[i][0]) }, style: {'background-color': 'red', 'text-background-color':'#FFFFFF', 'text-background-opacity': 0.8, 'font-size': 16}};
    let nodeCE = { data: { id: conceitoEfeito[i][0] }, position: { x: parseInt(posXConceitoEfeito[i][0]), y: parseInt(posYConceitoEfeito[i][0]) }, style: {'background-color': 'blue', 'text-background-color':'#FFFFFF', 'text-background-opacity': 0.8, 'font-size': 16} };
    let edge = { data: { id: i, source: conceitoCausa[i][0], target: conceitoEfeito[i][0]}, style: {'label': valorSinal[i], 'text-background-color':'#FFFFFF', 'text-background-opacity': 1, 'font-size': 20} };
    elements.push(nodeCC);
    elements.push(nodeCE);
    elements.push(edge);
  } 
  
  var dataFinal = JSON.stringify(elements);
  var jsonData = JSON.parse(dataFinal);

  return jsonData;
}