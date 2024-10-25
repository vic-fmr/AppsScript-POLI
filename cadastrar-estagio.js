function acessarPlanilha() {
  let planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let planilhaDestino = SpreadsheetApp.openById('1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0').getSheetByName('CONTROLE - 2022.2');

  coletarValores(planilhaAtiva, planilhaDestino)
}

function coletarValores(planilhaAtiva, planilhaDestino){
  
let numeroLinha = planilhaAtiva.getActiveCell().getRow();
let ultimaColuna = planilhaAtiva.getLastColumn();
let informacoesLinhaRange = planilhaAtiva.getRange(numeroLinha, 1, 1, ultimaColuna);
let informacoesLinha = informacoesLinhaRange.getValues();
let ultimaLinhaVaziaDestino = planilhaDestino.getLastRow() + 1;
let rangeDestino = planilhaDestino.getRange(ultimaLinhaVaziaDestino, 2, 1, informacoesLinha[0].length);

cadastrarValores(rangeDestino)
}

function cadastrarValores(rangeDestino){
rangeDestino.setValues(informacoesLinha);
let pintarVermelho = planilhaDestino.getRange(ultimaLinhavazia, 1, 1, planilhaDestino.getLastColumn());
pintarVermelho.setBackground('red');
informacoesLinhaRange.setBackground('red');
}


