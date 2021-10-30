function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function pegador(){
  const dataSheet = SpreadsheetApp.getActive().getSheetByName("Dados - Mapa Conceitual");
  const elements = [];
  const lastRow = dataSheet.getRange(1, 11).getValue();
  const conceitoCausa = dataSheet.getRange(1, 1, lastRow, 1).getValues();
  const conceitoEfeito = dataSheet.getRange(1, 2, lastRow, 1).getValues();
  const posXConceitoCausa = dataSheet.getRange(1, 3, lastRow, 1).getValues();
  const posYConceitoCausa = dataSheet.getRange(1, 4, lastRow, 1).getValues();
  const posXConceitoEfeito = dataSheet.getRange(1, 5, lastRow, 1).getValues();
  const posYConceitoEfeito = dataSheet.getRange(1, 6, lastRow, 1).getValues();
  const valorSinal = dataSheet.getRange(1, 7, lastRow, 1).getValues();

  for(let i = 0; i < lastRow; i++){
    let nodeCC = { data: { id: conceitoCausa[i][0] }, position: { x: parseInt(posXConceitoCausa[i][0]), y: parseInt(posYConceitoCausa[i][0]) }, style: {'background-color': 'red', 'text-background-color':'#FFFFFF', 'text-background-opacity': 0.2, 'font-size': 16}};
    let nodeCE = { data: { id: conceitoEfeito[i][0] }, position: { x: parseInt(posXConceitoEfeito[i][0]), y: parseInt(posYConceitoEfeito[i][0]) }, style: {'background-color': 'blue', 'text-background-color':'#FFFFFF', 'text-background-opacity': 0.2, 'font-size': 16} };
    let edge = { data: { id: i, source: conceitoEfeito[i][0], target: conceitoCausa[i][0]}, style: {'label': valorSinal[i], 'text-background-color':'#FFFFFF', 'text-background-opacity': 1, 'font-size': 20} };
    elements.push(nodeCC);
    elements.push(nodeCE);
    elements.push(edge);
  } 
  
  var dataFinal = JSON.stringify(elements);
  var jsonData = JSON.parse(dataFinal);
  return jsonData;
}