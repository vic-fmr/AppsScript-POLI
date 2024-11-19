function acessarPlanilha() {
  const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const planilhaDestino = SpreadsheetApp.openById('1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0').getSheetByName('CONTROLE - 2022.2');

  coletarValores(planilhaAtiva, planilhaDestino)
}

function coletarValores(planilhaAtiva, planilhaDestino) {
  const numeroLinha = planilhaAtiva.getActiveCell().getRow();
  const ultimaColuna = planilhaAtiva.getLastColumn();
  const informacoesLinhaRange = planilhaAtiva.getRange(numeroLinha, 1, 1, ultimaColuna);
  const informacoesLinha = informacoesLinhaRange.getValues();
  const ultimaLinhaVaziaDestino = planilhaDestino.getLastRow() + 1;
  const ultimaLinhaDestino = planilhaDestino.getLastRow()
  const rangeDestino = planilhaDestino.getRange(ultimaLinhaVaziaDestino, 2, 1, informacoesLinha[0].length);


  verificarCpf(informacoesLinha, ultimaLinhaDestino, planilhaDestino, rangeDestino, ultimaLinhaVaziaDestino, informacoesLinhaRange)
}

function verificarCpf(informacoesLinha, ultimaLinhaDestino, planilhaDestino, rangeDestino, ultimaLinhaVaziaDestino, informacoesLinhaRange) {

  cpf = informacoesLinha[0][3]
  cpfsCadastrados = planilhaDestino.getRange(9, 5, ultimaLinhaDestino, 1).getValues()

  dataInicio = new Date(informacoesLinha[0][10])
  isCadastrada = false;


  for (i = 0; i <= ultimaLinhaDestino; i++) {
    if (cpfsCadastrados[i] == cpf) {
      dataFimEstagio = new Date(planilhaDestino.getRange(i + 9, 13).getValue())
      Logger.log(dataFimEstagio)
      if (dataFimEstagio > dataInicio) {
        isCadastrada = true;
        break;
      }
    }
  }
  if (isCadastrada) {
    SpreadsheetApp.getUi().alert(`Aluno com cpf ${cpf} já cadastrado e com estágio em andamento, por favor verificar situação.`)
    return
  }
  cadastrarValores(rangeDestino, informacoesLinha, planilhaDestino, ultimaLinhaVaziaDestino, informacoesLinhaRange)

}


function cadastrarValores(rangeDestino, informacoesLinha, planilhaDestino, ultimaLinhaVaziaDestino, informacoesLinhaRange) {
  rangeDestino.setValues(informacoesLinha);
  const pintarVermelho = planilhaDestino.getRange(ultimaLinhaVaziaDestino, 1, 1, planilhaDestino.getLastColumn());
  pintarVermelho.setBackground('red');
  informacoesLinhaRange.setBackground('red');

}

