let ssDados = SpreadsheetApp.openByUrl('url da planilha completa');
let abaDADOS = ssDados.getSheetByName('nome da aba');

function rodarResultados() {
  clearFiles()

  let dados = abaDADOS.getRange('B2:V').getValues().filter(function(item) {return item[0] !=""});
//22
  for(i = 0;i < dados.length; i++){
    alterarTemplate(dados[i][0],dados[i][1],dados[i][2],dados[i][3],dados[i][4],dados[i][5],dados[i][6],dados[i][7],dados[i][8],dados[i][9],dados[i][10],dados[i][11],dados[i][12],dados[i][13],dados[i][14],dados[i][15], dados[i][16], dados[i][17], dados[i][18], dados[i][19], dados[i][20], dados[i][21], dados[i][22])
  }
}

function alterarTemplate(A,T,B, C, D, E, F, G, U, H, I, J, K, L, M, N, O, P, Q, R, S) {
  let pastaDestino = DriveApp.getFolderById('pasta do google(só o id sem o inicio)');
  let nomeArquivo = A + " - " + B 

  let idTemplate = 'id do template do google docs(sem inicio)';
  let novoArquivo = DriveApp.getFileById(idTemplate).makeCopy(nomeArquivo,destination=pastaDestino);
  let novoDoc = DocumentApp.openById(novoArquivo.getId());
  let docCorpo = novoDoc.getBody();

  docCorpo.replaceText("{A}", A);
  docCorpo.replaceText("{T}", T);
  docCorpo.replaceText("{B}", B);
  docCorpo.replaceText("{C}", C);
  docCorpo.replaceText("{D}", D);
  docCorpo.replaceText("{E}", E);
  docCorpo.replaceText("{F}", F);
  docCorpo.replaceText("{G}", G);
  docCorpo.replaceText("{U}", U);
  docCorpo.replaceText("{H}", H);
  docCorpo.replaceText("{I}", I);
  docCorpo.replaceText("{J}", J);
  docCorpo.replaceText("{K}", K);
  docCorpo.replaceText("{L}", L);
  docCorpo.replaceText("{M}", M);
  docCorpo.replaceText("{N}", N);
  docCorpo.replaceText("{O}", O);
  docCorpo.replaceText("{P}", P);
  docCorpo.replaceText("{Q}", Q);
  docCorpo.replaceText("{R}", R);
  docCorpo.replaceText("{S}", S);

  novoDoc.saveAndClose();

  let pdfBlob = novoDoc.getAs(MimeType.PDF);
  let pdf = pastaDestino.createFile(pdfBlob).setName(nomeArquivo);
  DriveApp.getFileById(novoDoc.getId()).setTrashed(true);


}

function onOpen(e) {
  let app = SpreadsheetApp;
  let ui = app.getUi();

  ui.createMenu("Ações")
  .addItem("Exportar para PDF", "rodarResultados")
  .addToUi();
}

function clearFiles(){
  let files = DriveApp.getFolderById('pasta do google(só o id sem o inicio)').getFiles();
  while(files.hasNext()){
    files.next().setTrashed(true);
  }
}
