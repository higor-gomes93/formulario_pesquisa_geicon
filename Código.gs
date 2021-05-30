function transformacaoAtuacao() {
  // Definindo as constantes
  const dataset = SpreadsheetApp.getActive().getSheetByName("atuacao_1");
  const extendedDataset = SpreadsheetApp.getActive().getSheetByName("atuacao_2");
  
  // Iterando
  for(let i = 0; i < 10; i++){
    let answers = dataset.getRange(2, 2+i, 207, 1).getValues();
    let constant = dataset.getRange(2, 1, 207, 1).getValues();
    let lastrow = extendedDataset.getLastRow() + 1;
    extendedDataset.getRange(lastrow, 1, 207, 1).setValues(constant);
    extendedDataset.getRange(lastrow, 2, 207, 1).setValues(answers);
  }
}


function transformacaoDedicacao() {
  // Definindo as constantes
  const dataset = SpreadsheetApp.getActive().getSheetByName("dedicacao_1");
  const extendedDataset = SpreadsheetApp.getActive().getSheetByName("dedicacao_2");
  
  // Iterando
  for(let i = 0; i < 3; i++){
    let answers = dataset.getRange(2, 2+i, 207, 1).getValues();
    let constant = dataset.getRange(2, 1, 207, 1).getValues();
    let lastrow = extendedDataset.getLastRow() + 1;
    extendedDataset.getRange(lastrow, 1, 207, 1).setValues(constant);
    extendedDataset.getRange(lastrow, 2, 207, 1).setValues(answers);
  }
}


function transformacaoCargo() {
  // Definindo as constantes
  const dataset = SpreadsheetApp.getActive().getSheetByName("cargo_1");
  const extendedDataset = SpreadsheetApp.getActive().getSheetByName("cargo_2");
  
  // Iterando
  for(let i = 0; i < 12; i++){
    let answers = dataset.getRange(2, 2+i, 207, 1).getValues();
    let constant = dataset.getRange(2, 1, 207, 1).getValues();
    let lastrow = extendedDataset.getLastRow() + 1;
    extendedDataset.getRange(lastrow, 1, 207, 1).setValues(constant);
    extendedDataset.getRange(lastrow, 2, 207, 1).setValues(answers);
  }
}


function transformacaoContribuicao() {
  // Definindo as constantes
  const dataset = SpreadsheetApp.getActive().getSheetByName("contribuicao_1");
  const extendedDataset = SpreadsheetApp.getActive().getSheetByName("contribuicao_2");
  
  // Iterando
  for(let i = 0; i < 2; i++){
    let answers = dataset.getRange(2, 2+i, 207, 1).getValues();
    let constant = dataset.getRange(2, 1, 207, 1).getValues();
    let lastrow = extendedDataset.getLastRow() + 1;
    extendedDataset.getRange(lastrow, 1, 207, 1).setValues(constant);
    extendedDataset.getRange(lastrow, 2, 207, 1).setValues(answers);
  }
}


function transformacaoOcupacao() {
  // Definindo as constantes
  const dataset = SpreadsheetApp.getActive().getSheetByName("ocupacao_1");
  const extendedDataset = SpreadsheetApp.getActive().getSheetByName("ocupacao_2");
  
  // Iterando
  for(let i = 0; i < 2; i++){
    let answers = dataset.getRange(2, 2+i, 207, 1).getValues();
    let constant = dataset.getRange(2, 1, 207, 1).getValues();
    let lastrow = extendedDataset.getLastRow() + 1;
    extendedDataset.getRange(lastrow, 1, 207, 1).setValues(constant);
    extendedDataset.getRange(lastrow, 2, 207, 1).setValues(answers);
  }
}























