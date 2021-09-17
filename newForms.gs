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
  const proxIter = formulario.getRange(20,5).getValue();

  // Definindo os estilos da próxima iteração
  const camposIncluir = auxiliares.getRange(2, 19, 11, 1);
  const camposLigar = auxiliares.getRange(2, 20, 5, 1);

  // Ajustando para a próxima iteração
  formulario.getRange(10, 4).setValue(proxIter);
  formulario.getRange(11, 5, 13, 1).clear();
  if(proxIter == "Incluir um novo conceito"){
    camposIncluir.copyTo(formulario.getRange(11, 5, 11, 1));
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IF($D$11:$D$23="";TRUE;FALSE)').setRanges([formulario.getRange(11, 4, 13, 2)]).build();
    SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setRanges([formulario.getRange(11, 5)]).build();
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=IF(AND($E$13=FALSE;$E$14=FALSE);TRUE;FALSE)').setRanges([formulario.getRange(13, 5, 2, 1)]).build();
    SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setRanges([formulario.getRange(15, 5, 2, 1)]).build();
    SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setRanges([formulario.getRange(18, 5, 4, 1)]).build();
  } else if(proxIter == "Apenas conectar conceitos"){
    camposLigar.copyTo(formulario.getRange(11, 5, 5, 1));
  }

}