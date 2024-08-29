function cadastro() {
  // Obter a planilha ativa e a célula ativa
  var planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celulaAtiva = planilhaAtiva.getActiveCell();
  
  // Obter a planilha de destino (substitua o ID e nome pela sua planilha e aba)
  var planilhaDestino = SpreadsheetApp.openById('1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0').getSheetByName('CONTROLE - 2022.2');
  
  // Obter o número da linha da célula ativa
  var numeroLinha = celulaAtiva.getRow();
  
  // Selecionar toda a linha onde está a célula ativa
  var linhaParaCopiar = planilhaAtiva.getRange(numeroLinha, 1, 1, planilhaAtiva.getLastColumn()).getValues();
  var restoParaCopiar = planilhaDestino.getRange(planilhaDestino.getLastRow(), planilhaAtiva.getLastColumn() + 1, 1, planilhaDestino.getLastColumn());
  
  // Seleciona o resto da linha na planinha destino, visto que o numero de colunas na linha da planilha original é menor
  var destinoResto = planilhaDestino.getRange(planilhaDestino.getLastRow() + 1, planilhaAtiva.getLastColumn() + 1, 1, planilhaDestino.getLastColumn());
  var resto = restoParaCopiar.getFormulas();

  // Obter a última linha vazia na planilha de destino
  var ultimaLinhavazia = planilhaDestino.getLastRow() + 1;

  // Colar a linha selecionada na planilha de destino
  var rangeDestino = planilhaDestino.getRange(ultimaLinhavazia, 2, 1, linhaParaCopiar[0].length);
  rangeDestino.setValues(linhaParaCopiar);
  destinoResto.setFormulas(resto);
  
  //Pinta a linha destino de vermelho
  var pintarVermelho = planilhaDestino.getRange(ultimaLinhavazia, 1, 1, planilhaDestino.getLastColumn());
  pintarVermelho.setBackground('red');

 
  //Formata a célula presente na coluna com os checkbox
  var caixinha = planilhaDestino.getRange(ultimaLinhavazia, 45)

  var rule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();

      caixinha.setDataValidation(rule);
}


