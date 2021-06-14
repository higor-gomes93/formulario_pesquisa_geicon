function coletorPastas() {
  // Definindo a pasta de respostas
  const folder = DriveApp.getFolderById("1UWh2jsB0_lxkovCbETolUhN3dYETDQ3I");  // Pasta central de respostas
  const folderIterator = folder.getFolders();
  let folderCheck = folderIterator.hasNext();
  const folders = new Array();  // Array com as pastas de cada turma

  // Coletando todas as subpastas com as respostas
  while (folderCheck == true){
    let folderId = folderIterator.next().getId().toString();
    folders.push(folderId);
    folderCheck = folderIterator.hasNext();
  }

  // Retornando os valores
  return folders
}

function coletorArquivos(folderId) {
  // Definindo a pasta de respostas
  const folder = DriveApp.getFolderById(folderId);
  const fileterator = folder.getFiles();
  let fileCheck = fileterator.hasNext();
  const files = new Array();

  // Coletando todas as subpastas com as respostas
  while (fileCheck == true){
    let fileId = fileterator.next().getId().toString();
    files.push(fileId);
    fileCheck = fileterator.hasNext();
  }

  // Retornando os valores
  return files
}

function coletorDados(fileId) {
  // Definindo a origem dos dados
  const spreadsheetId = fileId;
  const sheetName = "Dados";
  const startRow = 2;  
  const startColumn = 1;
  const endRow = 6;
  const endColumn = 6; 
  
  // Coletando os valores
  const data = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName).getRange(startRow, startColumn, endRow-1, endColumn).getValues();

  // Retornando os valores
  return data
}

function coletorNomes(fileId) {
  // Coletando o nome da planilha
  const rawName = DriveApp.getFileById(fileId).getName().toString();
  const name = rawName.split(" - ")[1];
  const listNames = [name, name, name, name, name];

  // Retornando o valor
  return listNames
}

function registroDados(data, names) {
   // Definindo a planilha destino
  const mainSpreadsheetId = "1RtPiORyQqR1eyq7_n8BPuLK9JeEX3FgVKBfGfM0_ns8";
  const mainSheetName = "Dataset";
  const mainSheet = SpreadsheetApp.openById(mainSpreadsheetId).getSheetByName(mainSheetName);
  const mainstartRow = mainSheet.getLastRow();
  const mainstartColumn = 1;
  const mainendRow = 5;
  const mainendColumn = 6;
  
  // Inserindo os dados
  mainSheet.getRange(mainstartRow+1, mainstartColumn+1, mainendRow, mainendColumn).setValues(data);
  mainSheet.getRange(mainstartRow+1, mainstartColumn, mainendRow, 1).setValue(names);
}

function coletorRespostas() {
  const folders = coletorPastas();

  for (let i in folders) {
    files = coletorArquivos(folders[i]);
    for (let j in files) {
      dados = coletorDados(files[j]);
      nome = coletorNomes(files[j]);
      registroDados(dados, nome);
    }
  }
}