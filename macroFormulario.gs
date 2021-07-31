function salvarRespostas() {
  // Definindo as abas
  const formulario = SpreadsheetApp.getActive().getSheetByName("Formulário");
  const auxiliares = SpreadsheetApp.getActive().getSheetByName("Tabelas Auxiliares");
  const respostas = SpreadsheetApp.getActive().getSheetByName("Respostas");
  const textoSheet = SpreadsheetApp.getActive().getSheetByName("Texto");
  const relatorio = SpreadsheetApp.getActive().getSheetByName("Relatório");

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
  const data = Utilities.formatDate(new Date(),"GMT-3", "dd/MM/yyyy");

  // Conferindo as respostas
  // Criando os elementos do texto
  let elementos = [];
  elementos = [
    pergunta1.getValue() + ' - ' + conceitoCausaOrigem2.getValue() + '\n',
    pergunta2.getValue() + ' - ' + camadaUmOrigem.getValue() + '\n',
    pergunta3.getValue() + ' - ' + camadaDoisOrigem.getValue() + '\n',
    pergunta4.getValue() + ' - ' + camadaTresOrigem.getValue() + '\n',
    pergunta5.getValue() + ' - ' + novoConceitoEfeitoOrigem.getValue() + '\n'
  ];
  
  // Criando o texto
  const texto = elementos.join('\n');

  // UI
  const resposta = SpreadsheetApp.getUi().alert("Confira suas respostas.", texto, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

  if(resposta == SpreadsheetApp.getUi().Button.CANCEL){
    SpreadsheetApp.getUi().alert("Cancelado! Repita o processo.", SpreadsheetApp.getUi().ButtonSet.OK);
    return
  }

  // Encontrando a última linha na aba Respostas
  const linha1 = respostas.getRange(1, 15).getValue();

  // Definindo as variáveis destino
  const conceitoCausaDestino = respostas.getRange(linha1, 2);
  const conceitoEfeitoDestino = respostas.getRange(linha1, 3);
  const camadaUmDestino = respostas.getRange(linha1, 4);
  const camadaDoisDestino = respostas.getRange(linha1, 5);
  const camadaTresDestino = respostas.getRange(linha1, 6);
  const dataColeta = respostas.getRange(linha1, 13);
  const novoConceitoEfeitoDestino = auxiliares.getRange(2, 2);
  const dataMce = relatorio.getRange(9, 8);
  const dataEA = relatorio.getRange(25, 8);
  const dataTP = relatorio.getRange(45, 8);

  
  // Inserindo os valores
  conceitoCausaDestino.setValue(conceitoCausaOrigem2.getDisplayValue());
  conceitoEfeitoDestino.setValue(conceitoEfeitoOrigem.getValue());
  camadaUmDestino.setValue(camadaUmOrigem.getValue());
  camadaDoisDestino.setValue(camadaDoisOrigem.getValue());
  camadaTresDestino.setValue(camadaTresOrigem.getValue());
  novoConceitoEfeitoDestino.setValue(novoConceitoEfeitoOrigem.getValue());
  dataColeta.setValue(data);
  dataMce.setValue(data);
  dataEA.setValue(data);
  dataTP.setValue(data);

  // Inserindo valores textuais
  const textLastRow = textoSheet.getLastRow();
  textoSheet.getRange(textLastRow+1, 1).setValue(pergunta1.getValue());
  textoSheet.getRange(textLastRow+1, 2).setValue(conceitoCausaOrigem.getValue());
  textoSheet.getRange(textLastRow+1, 3).setValue(pergunta2.getValue());
  textoSheet.getRange(textLastRow+1, 4).setValue(camadaUmOrigem.getValue());
  textoSheet.getRange(textLastRow+1, 5).setValue(pergunta3.getValue());
  textoSheet.getRange(textLastRow+1, 6).setValue(camadaDoisOrigem.getValue());
  textoSheet.getRange(textLastRow+1, 7).setValue(pergunta4.getValue());
  textoSheet.getRange(textLastRow+1, 8).setValue(camadaTresOrigem.getValue());  
  textoSheet.getRange(textLastRow+1, 9).setValue(data);

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
  
  // Log de confirmação
  SpreadsheetApp.getUi().alert("Concluído!", "Respostas salvas com sucesso.", SpreadsheetApp.getUi().ButtonSet.OK);
}
