function sendEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Envio de E-mails");    
  const dataRange = sheet.getDataRange();            
  const data = dataRange.getValues();           
  
  for (let i=2; i<data.length;i++){  
    let row = data[i];
    let email_destinatario = row[2];                                                       
    let subject = row[5];                                                                  
    let message = row[6];
    let check = row[7];
    let planilha = row[3].toString();                                                               
    let options = {htmlBody: message, cc: "h146435@dac.unicamp.br"};                                             

    if(check != "ENVIADO" & subject != ""){
      const id = SpreadsheetApp.openByUrl(planilha).getId();
      DriveApp.getFileById(id).addEditor(email_destinatario);
      MailApp.sendEmail(email_destinatario, subject, message, options);
      sheet.getRange(i+1, 8).setValue("ENVIADO");   
      SpreadsheetApp.flush();                                       
    }
  }
}