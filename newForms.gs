function inicioElicitacao() {
  // Definição das abas
  const formulario = SpreadsheetApp.getActive().getSheetByName("Formulário");
  const auxiliares = SpreadsheetApp.getActive().getSheetByName("Auxiliares");
  const respostas = SpreadsheetApp.getActive().getSheetByName("Respostas");

  // Coletando os conceitos
  const conceitoUm = auxiliares.getRange(13,10).getValue();
  const conceitoDois = auxiliares.getRange(14,10).getValue();
  const posicaoUm = auxiliares.getRange(13,15).getValue();
  const posicaoDois = auxiliares.getRange(14,15).getValue();
  const tipoUm = auxiliares.getRange(13,11).getValue();

  // Encontrando as posições na aba de respostas
  const lastRow = respostas.getRange(1, 14).getValue();
  const conceitoCC = respostas.getRange(lastRow, 2);
  const conceitoCE = respostas.getRange(lastRow, 4);
  const posicaoCC = respostas.getRange(lastRow, 3);
  const posicaoCE = respostas.getRange(lastRow, 5);
  
  // Inserindo os conceitos
  if(tipoUm == "Efeito"){
    conceitoCC.setValue(conceitoUm);
    conceitoCE.setValue(conceitoDois);
    posicaoCC.setValue(posicaoUm);
    posicaoCE.setValue(posicaoDois);
  } else if(tipoUm == "Causa"){
    conceitoCC.setValue(conceitoDois);
    conceitoCE.setValue(conceitoUm);
    posicaoCC.setValue(posicaoDois);
    posicaoCE.setValue(posicaoUm);
  } else {
    console.log("Erro");
  }

  // Coletando o tipo da próxima iteração
  const proxIter = formulario.getRange(21,5).getValue();

  // Definindo os estilos da próxima iteração
  const camposIncluir = auxiliares.getRange(2, 19, 11, 1);
  const camposLigar = auxiliares.getRange(2, 20, 5, 1);

  // Ajustando para a próxima iteração
  formulario.getRange(11, 4).setValue(proxIter);
  formulario.getRange(12, 5, 13, 1).clear();
  if(proxIter == "Incluir um novo conceito"){
    // Copiando e colando os campos
    camposIncluir.copyTo(formulario.getRange(12, 5, 11, 1));
    // Criando as formatações condicionais
    const rule1 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IF($D$11:$D$23="";TRUE;FALSE)').setBackground('#ffffff').setRanges([formulario.getRange(12, 4, 13, 2)]).build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#fff2cc').setRanges([formulario.getRange(12, 5)]).build();
    const rule3 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IF(AND($E$13=FALSE;$E$14=FALSE);TRUE;FALSE)').setBackground('#fff2cc').setRanges([formulario.getRange(14, 5, 2, 1)]).build();
    const rule4 = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#fff2cc').setRanges([formulario.getRange(16, 5, 2, 1)]).build();
    const rule5 = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#fff2cc').setRanges([formulario.getRange(19, 5, 4, 1)]).build();
    // Criando o vetor de formatações
    const rules = formulario.getConditionalFormatRules();
    rules.push(rule1);
    rules.push(rule2);
    rules.push(rule3);
    rules.push(rule4);
    rules.push(rule5);
    // Inserindo as formatações
    formulario.setConditionalFormatRules(rules);
    // Inserindo a formatação de dados
    formulario.getRange(19, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(auxiliares.getRange(30, 23, 100, 1)).setAllowInvalid(false).setHelpText("Escolha um dos conceitos da lista.").build());
    formulario.getRange(12, 5, 13, 1).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.DASHED);
  } else if(proxIter == "Apenas conectar conceitos"){
    formulario.getRange(12, 5, 13, 1).clearDataValidations();
    // Copiando e colando os campos
    camposLigar.copyTo(formulario.getRange(12, 5, 5, 1));
    // Criando as formatações condicionais
    const rule1 = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IF($D$11:$D$23="";TRUE;FALSE)').setBackground('#ffffff').setRanges([formulario.getRange(12, 4, 13, 2)]).build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground('#fff2cc').setRanges([formulario.getRange(12, 5, 5, 1)]).build();
    // Criando o vetor de formatações
    const rules = formulario.getConditionalFormatRules();
    rules.push(rule1);
    rules.push(rule2);
    // Inserindo as formatações
    formulario.setConditionalFormatRules(rules);
    // Inserindo a formatação de dados
    formulario.getRange(12, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(auxiliares.getRange(30, 23, 100, 1)).setAllowInvalid(false).setHelpText("Escolha um dos conceitos da lista.").build());
    formulario.getRange(13, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(auxiliares.getRange(30, 23, 100, 1)).setAllowInvalid(false).setHelpText("Escolha um dos conceitos da lista.").build());
    formulario.getRange(12, 5, 13, 1).setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.DASHED);
  }
}