function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}


function pegador(){
  var elements = [];
  var conceitoCausa = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 2, 5, 1).getValues();
  var conceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 3, 5, 1).getValues();
  var posXConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 4, 5, 1).getValues();
  var posYConceitoCausa = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 5, 5, 1).getValues();
  var posXConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 6, 5, 1).getValues();
  var posYConceitoEfeito = SpreadsheetApp.getActive().getSheetByName("Teste").getRange(1, 7, 5, 1).getValues();

  for(let i = 0; i < 5; i++){
    let nodeCC = { data: { id: conceitoCausa[i][0] }, position: { x: parseInt(posXConceitoCausa[i][0]), y: parseInt(posYConceitoCausa[i][0]) } };
    let nodeCE = { data: { id: conceitoEfeito[i][0] }, position: { x: parseInt(posXConceitoEfeito[i][0]), y: parseInt(posYConceitoEfeito[i][0]) } }
    let edge = { data: { id: i, source: conceitoCausa[i][0], target: conceitoEfeito[i][0]} };
    elements.push(nodeCC);
    elements.push(nodeCE);
    elements.push(edge);
  } 
  /*
  var elements = [
    // nodes
    { data: { id: 'a' }, position: { x: 100, y: 100 } },
    { data: { id: 'b' }, position: { x: 200, y: 200 } },
    { data: { id: 'c' }, position: { x: 300, y: 300 } },
    { data: { id: 'd' }, position: { x: 400, y: 400 } },
    { data: { id: 'e' }, position: { x: 500, y: 500 } },
    { data: { id: 'f' }, position: { x: 600, y: 600 } },
    { data: { id: 'b' }, position: { x: 700, y: 700 } },
    // edges
    {
      data: {
        id: 'ab',
        source: 'a',
        target: 'b'
      }
    },
    {
      data: {
        id: 'cd',
        source: 'c',
        target: 'd'
      }
    },
    {
      data: {
        id: 'ef',
        source: 'e',
        target: 'f'
      }
    },
    {
      data: {
        id: 'ac',
        source: 'a',
        target: 'c'
      }
    },
    {
      data: {
        id: 'be',
        source: 'b',
        target: 'e'
      }
    },
    {
      data: {
        id: 'ba',
        source: 'b',
        target: 'a'
      }
    }
  ];
  */


  var dataFinal = JSON.stringify(elements);
  var jsonData = JSON.parse(dataFinal);

  return jsonData;
}