function geradorFormularios() {
  // Definindo as planilhas e as pastas
  const planilha_controle = SpreadsheetApp.openById("1GrJuxWZ3cqFveAGZ7rnqqI2KGgW0Ubb_l736pUAg7Oc");  // Planilha para envio de e-mails e geração de formulários
  const formularios = DriveApp.getFolderById("1GUrCWml7td6lHr8v9C-tq0jOaS7TUbt0");  // Pasta contendo todos os formulários

  // Validar se existe algum dado

  // Copiando os formulários e inserindo os links
  const linhas = planilha_controle.getSheetByName("Envio de E-mails").getLastRow();
  for(let i = 3; i <= linhas; i++){
    let formulario_pessoal = DriveApp.getFileById("18bBvISV9pC4UnSnzUDSSb88XwuIySm6O7X9gynkfx8c").makeCopy().moveTo(formularios);
    formulario_pessoal.setName("Formulário - "+planilha_controle.getSheetByName("Envio de E-mails").getRange(i, 1).getValue());
    SpreadsheetApp.open(formulario_pessoal).getSheetByName("Tabelas Auxiliares").getRange(1, 21).setValue(planilha_controle.getSheetByName("Envio de E-mails").getRange(i, 1).getValue())
    planilha_controle.getSheetByName("Envio de E-mails").getRange(i, 4).setValue(formulario_pessoal.getUrl());
  }

  console.log(linhas);

  // Arrastando as fórmulas
  planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 2).autoFill(planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 2, linhas-2, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 5).autoFill(planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 5, linhas-2, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 6).autoFill(planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 6, linhas-2, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 7).autoFill(planilha_controle.getSheetByName("Envio de E-mails").getRange(3, 7, linhas-2, 1), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // Alerta
  SpreadsheetApp.getUi().alert("Ferramenta configurada.", SpreadsheetApp.getUi().ButtonSet.OK);
}
