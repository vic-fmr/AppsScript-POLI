/** @NotOnlyCurrentDoc */

//Selecionar planilha e linha ativa. Gera uma array para os links
function abrirDocumentos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getActiveSheet();
  var linha = aba.getActiveRange().getRow();
  var links = [];

//Selecionar cada link
  for (var coluna = 22; coluna <= 25; coluna++) {
    var linkDoc = aba.getRange(linha, coluna).getValue();

    //Separar em caso de virgulas 
    if (linkDoc) {
      var linkDocVirgula = linkDoc.split(',');
      linkDocVirgula.forEach(function(link) {
        links.push(link.trim());
      });
    }
    }
  
  for (var coluna = 27; coluna <= 30; coluna++) {
    var linkDoc = aba.getRange(linha, coluna).getValue();
    if (linkDoc) {
      var linkDocVirgula = linkDoc.split(',');
      linkDocVirgula.forEach(function(link) {
        links.push(link.trim());
      });
    }
    }
  
//Abrir os documentos no navegador
  var page = '<script>';
  links.forEach(function(link) {
    page += 'window.open("' + link + '");';
  });
  page += 'google.script.host.close();</script>';
 
 //Caixa de diálogo
  var interface = HtmlService
    .createHtmlOutput(page)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(1)
    .setHeight(1);   
  SpreadsheetApp.getUi().showModalDialog(interface, 'Abrindo documentos...');

}
