function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function pegador(){
  var elements = [];
  var conceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 2, 5, 1).getValues();
  var conceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 3, 5, 1).getValues();
  var posXConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 5, 5, 1).getValues();
  var posYConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 4, 5, 1).getValues();
  var posXConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 7, 5, 1).getValues();
  var posYConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Dataset_Filtrado").getRange(1, 6, 5, 1).getValues();

  for(let i = 0; i < 5; i++){
    let nodeCC = { data: { id: conceitoCausa[i][0] }, position: { x: parseInt(posXConceitoCausa[i][0]), y: parseInt(posYConceitoCausa[i][0]) } };
    let nodeCE = { data: { id: conceitoEfeito[i][0] }, position: { x: parseInt(posXConceitoEfeito[i][0]), y: parseInt(posYConceitoEfeito[i][0]) } }
    let edge = { data: { id: i, source: conceitoCausa[i][0], target: conceitoEfeito[i][0]} };
    elements.push(nodeCC);
    elements.push(nodeCE);
    elements.push(edge);
  } 
  
  var dataFinal = JSON.stringify(elements);
  var jsonData = JSON.parse(dataFinal);

  return jsonData;
}