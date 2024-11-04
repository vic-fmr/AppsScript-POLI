function abrirDialogo() {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('Interface Notas')
        .setWidth(700)
        .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Insira as Notas');
  }
  
  function receberDados(nota_supervisorValor, nota_orientadorValor) {
    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let linhaAtiva = planilha.getActiveCell().getRow();
    let cpfAluno = planilha.getRange(linhaAtiva, 7).getValue();
    let linkRelatorio = planilha.getRange(linhaAtiva, 20).getValue(); 
  
  
    let planilhaDestino = SpreadsheetApp.openById("1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0").getActiveSheet();
    let ultimaLinha = planilhaDestino.getLastRow();

    let dadosDestino = planilhaDestino.getRange(3, 5, ultimaLinha - 2, 31).getValues();
  
    let linhaDestino = null;
    let contadorCPF = 0; 
    for (let i = 0; i < dadosDestino.length; i++) {
      let cpf_destino = dadosDestino[i][0]; 
      let tipoEstagio = dadosDestino[i][6]; 
  
 
      if (cpf_destino == cpfAluno && tipoEstagio === "Obrigatório") {
        contadorCPF++; 
  
        if (contadorCPF > 1) {
          
          SpreadsheetApp.getUi().alert('Erro: O CPF ' + cpfAluno + ' aparece mais de uma vez com estágio "Obrigatório". Verifique a planilha.');
          return; 
        }
  
        linhaDestino = i + 3; 
      }
    }
  
    if (linhaDestino) {
      planilhaDestino.getRange(linhaDestino, 33, 1, 3).setValues([[linkRelatorio, nota_supervisorValor, nota_orientadorValor]]);
    }
  }
  
