function salvarRespostas() {
  // Definindo as abas
  const formulario = SpreadsheetApp.getActive().getSheetByName("Formulário");
  const auxiliares = SpreadsheetApp.getActive().getSheetByName("Tabelas Auxiliares");
  const respostas = SpreadsheetApp.getActive().getSheetByName("Respostas");

  // Definindo as variáveis origem
  const conceitoCausaOrigem = formulario.getRange(9, 29);
  const conceitoCausaOrigem2 = auxiliares.getRange(3, 2);
  const conceitoEfeitoOrigem = auxiliares.getRange(2, 2);
  const camadaUmOrigem = formulario.getRange(10, 29);
  const camadaDoisOrigem = formulario.getRange(11, 29);
  const camadaTresOrigem = formulario.getRange(12, 29);
  const novoConceitoEfeitoOrigem = formulario.getRange(13, 29);
  const pergunta1 = formulario.getRange(9, 4);
  const pergunta2 = formulario.getRange(10, 4);
  const pergunta3 = formulario.getRange(11, 4);
  const pergunta4 = formulario.getRange(12, 4);
  const pergunta5 = formulario.getRange(13, 4);

  // Encontrando a porcentagem atual
  const log = auxiliares.getRange(1, 18).getValue();
  const novoLog = log+2;

  // Finalizando se já estiver concluído
  if(log == 10){
    SpreadsheetApp.getUi().alert("Você já finalizou o formulário.", SpreadsheetApp.getUi().ButtonSet.OK);
    return
  }

  // Conferindo as respostas
  // Criando os elementos do texto
  let elementos = [];
  if(log < 8){
    elementos = [
      pergunta1.getValue() + ' - ' + conceitoCausaOrigem2.getValue() + '\n',
      pergunta2.getValue() + ' - ' + camadaUmOrigem.getValue() + '\n',
      pergunta3.getValue() + ' - ' + camadaDoisOrigem.getValue() + '\n',
      pergunta4.getValue() + ' - ' + camadaTresOrigem.getValue() + '\n',
      pergunta5.getValue() + ' - ' + novoConceitoEfeitoOrigem.getValue() + '\n'
    ];
  } else if(log == 8){
    elementos = [
      pergunta1.getValue() + ' - ' + conceitoCausaOrigem2.getValue() + '\n',
      pergunta2.getValue() + ' - ' + camadaUmOrigem.getValue() + '\n',
      pergunta3.getValue() + ' - ' + camadaDoisOrigem.getValue() + '\n',
      pergunta4.getValue() + ' - ' + camadaTresOrigem.getValue() + '\n',
    ];
  }
  
  // PROMOÇÃO
  // Criando o texto
  const texto = elementos.join('\n');

  // UI
  const resposta = SpreadsheetApp.getUi().alert("Confira suas respostas.", texto, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

  if(resposta == SpreadsheetApp.getUi().Button.CANCEL){
    SpreadsheetApp.getUi().alert("Cancelado! Repita o processo.", SpreadsheetApp.getUi().ButtonSet.OK);
    return
  }

  // Encontrando a última linha na aba Respostas
  const linha1 = respostas.getRange(1, 14).getValue();

  // Definindo as variáveis destino
  const conceitoCausaDestino = respostas.getRange(linha1, 2);
  const conceitoEfeitoDestino = respostas.getRange(linha1, 3);
  const camadaUmDestino = respostas.getRange(linha1, 4);
  const camadaDoisDestino = respostas.getRange(linha1, 5);
  const camadaTresDestino = respostas.getRange(linha1, 6);
  const novoConceitoEfeitoDestino = auxiliares.getRange(2, 2);
  
  // Inserindo os valores
  conceitoCausaDestino.setValue(conceitoCausaOrigem2.getDisplayValue());
  conceitoEfeitoDestino.setValue(conceitoEfeitoOrigem.getValue());
  camadaUmDestino.setValue(camadaUmOrigem.getValue());
  camadaDoisDestino.setValue(camadaDoisOrigem.getValue());
  camadaTresDestino.setValue(camadaTresOrigem.getValue());
  novoConceitoEfeitoDestino.setValue(novoConceitoEfeitoOrigem.getValue());

  // Encontrando a última linha na aba Tabelas Auxiliares
  const linha2 = auxiliares.getRange(1, 15).getValue();

  // Definindo a variável de destino
  const local = auxiliares.getRange(linha2, 11);

  // Inserindo o conceito
  local.setValue(conceitoCausaOrigem2.getValue());

  // Limpando os valores
  conceitoCausaOrigem.clearContent();
  camadaUmOrigem.clearContent();
  camadaDoisOrigem.clearContent();
  camadaTresOrigem.clearContent();
  novoConceitoEfeitoOrigem.clearContent();
  
  // Inserindo os valores
  // Na barra de porcentagem
  if(log < 8){
    formulario.getRange(17, 18, 1, novoLog).setValue("*");
    formulario.getRange(17, 18+novoLog).setValue(novoLog+"0%").setFontColor("#666666");
  } else if(log == 8){
    formulario.getRange(16, 18, 1, novoLog).setValue("*");
    formulario.getRange(16, 18+novoLog).setValue(novoLog+"0%").setFontColor("#666666");
  }
  // Em Tabelas Auxiliares
  auxiliares.getRange(1, 18).setValue(novoLog);

  // Preparando para a rodada final
  if(novoLog == 8){
    formulario.deleteRow(13);
  } else if(novoLog == 10){
    auxiliares.getRange(3, 2).setValue("Finalizado!");
  }

  // Imprimindo alerta
  if(log <= 6){
    SpreadsheetApp.getUi().alert("Respostas salvas!", "Progresso atual: "+novoLog+"0%", SpreadsheetApp.getUi().ButtonSet.OK);
  } else if(log == 8){
    SpreadsheetApp.getUi().alert("Respostas salvas!", "Formulário concluído.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


function rodadaExtra() {
  // Definindo as abas
  const formulario = SpreadsheetApp.getActive().getSheetByName("Formulário");
  const auxiliares = SpreadsheetApp.getActive().getSheetByName("Tabelas Auxiliares");
  const respostas = SpreadsheetApp.getActive().getSheetByName("Respostas");

  // Criando o novo campo
  formulario.insertRowBefore(9);
  formulario.getRange(9, 3, 1, 26).merge();
  formulario.getRange(10, 3, 1, 26).copyFormatToRange(formulario.getRange(9, 3, 1, 26).getGridId(), 3, 28, 9, 9);
  formulario.getRange(10, 29, 1, 14).copyFormatToRange(formulario.getRange(9, 29, 1, 14).getGridId(), 29, 42, 9, 9);

  // Criando a lista interface
  const conceitosEfeito = auxiliares.getRange(4, 18, 5, 1);
  const conceitosCausa = auxiliares.getRange(4, 20, 5, 1);
  formulario.getRange(9, 29).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(conceitosEfeito).setAllowInvalid(false).build());
  auxiliares.getRange(4, 20).setFormula("=FILTER(Q4:Q;R4:R<>'Formulário 1'!AC9)");
  formulario.getRange(10, 29).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(conceitosCausa).setAllowInvalid(false).build());
  formulario.getRange(11, 29, 1, 14).copyFormatToRange(formulario.getRange(10, 29, 1, 14).getGridId(), 29, 42, 10, 10);
  const formula = auxiliares.getRange(4, 21).getFormula();
  formulario.getRange(10, 4).setFormula(formula);

  // Arrumando as categorias
  auxiliares.getRange(3, 2).clearContent();
  auxiliares.getRange(2, 2).setFormula("='Formulário 1'!AC9");
  auxiliares.getRange(3, 2).setFormula("='Formulário 1'!AC10");
}



function enviarFormulario(){
  
}

function macro() {
  const tipo1 = SpreadsheetApp.getActive().getSheetByName("Formulário").getRange(15, 3).getValue();
  const tipo2 = SpreadsheetApp.getActive().getSheetByName("Formulário").getRange(16, 3).getValue();

  // Rodando a função
  if (tipo1 == "Salvar Respostas" || tipo2 == "Salvar Respostas"){
    salvarRespostas();
  } else if (tipo1 == "Rodada Extra" || tipo2 == "Rodada Extra"){
    rodadaExtra();
  } else if (tipo1 == "Enviar Formulário" || tipo2 == "Enviar Formulário"){
    enviarFormulario();
  }
}
















