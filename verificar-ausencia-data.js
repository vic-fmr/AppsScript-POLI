// Script pra verificar cadastros que não possuem datas cadastradas (inicio ou fim)

function coletarCpfs() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet()
  let abaMatriz = planilha.getSheetByName("CONTROLE - 2022.2");
  let abaCpfsSemData = planilha.getSheetByName("Dados Para Corrigir na Matriz")

  let linhaInicial = abaCpfsSemData.getRange(1, 6).getValue();
  let ultimaLinha = abaMatriz.getLastRow() - (linhaInicial - 1);

  let datas = abaMatriz.getRange(linhaInicial, 12, ultimaLinha, 2).getValues()

  verficarExistenciaData(abaCpfsSemData, datas, abaMatriz)
}

function verficarExistenciaData(abaCpfsSemData, datas, abaMatriz) {
  let cpfsSemData = []

  for (i = 0; i < datas.length; i++) {
    if (i + 5 == 448) {
      continue
    }
    for (n = 0; n < datas[i].length; n++) {
      if (!datas[i][n]) {
        let cpfDaLinha = abaMatriz.getRange(i + 5, 5).getValue()

        if (!cpfDaLinha) {
          nomeDoAlunoSemCpf = abaMatriz.getRange(i + 5, 4).getValue()
          emailDoAlunoSemCpf = abaMatriz.getRange(i + 5, 7).getValue()

          cpfDaLinha = `CPF não cadastrado, Nome: ${nomeDoAlunoSemCpf} Email: ${emailDoAlunoSemCpf}`
        }

        cpfsSemData.push(cpfDaLinha)
      }
    }
  }
  guardarDadosNaAba(abaCpfsSemData, cpfsSemData)

}

function guardarDadosNaAba(abaCpfsSemData, cpfsSemData) {
  if (cpfsSemData) {

    let cpfsSemData_2D = cpfsSemData.map(cpf => [cpf]);
    abaCpfsSemData.getRange(4, 2, cpfsSemData_2D.length, 1).setValues(cpfsSemData_2D);
  } else {
    Logger.log("Nenhum cpf sem data preenchida")
  }
}
