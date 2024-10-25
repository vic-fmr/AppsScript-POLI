
function acessarPlanilha() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getActiveSheet();
  let linha = aba.getActiveRange().getRow();
  let links = [];

  coletarLinks(linha, links, aba)
}

//Bloco para capturar os links, alterar o valor de 'coluna' com base na planilha
function coletarLinks(linha, links, aba) {
  cellContent = []
  for (let coluna = 15; coluna <= 20; coluna++) {
    if (coluna == 17 || coluna == 19) { //Pular essas colunas
      continue;
    }
    cellContent.push(aba.getRange(linha, coluna).getValue());
  }
  tratarLinks(cellContent, links)
}
  
  //Bloco para separar o links por vírgulas e inserir em 'cellLinks" 
function tratarLinks(cellContent, links){
  cellLinks = [];
  cellContent.forEach(content => {
    if (content) {
      cellLinks.push(content.split(','))
    };
  });
  cellLinks.forEach((link) => {
    links.push(link);
  });

  abrirDocumentos(links);
}

//Bloco para criar o script dos documentos
function abrirDocumentos(links) {
  let page = '<script>';

  links.forEach((link) => {
    page += 'window.open("' + link + '");';
  });

  page += 'google.script.host.close();</script>';

  mostrarTela(page)
}

//Bloco para abrir os documentos nas abas
function mostrarTela(page) {
  let interface = HtmlService
    .createHtmlOutput(page)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(1)
    .setHeight(1);

  SpreadsheetApp.getUi().showModalDialog(interface, 'Abrindo documentos...');

  Logger.log(page)

}