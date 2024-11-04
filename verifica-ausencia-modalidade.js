function coletarCpfsModalidade() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet();
  let abaMatriz = planilha.getSheetByName("CONTROLE - 2022.2");
  let abacpfsSemModalidade = planilha.getSheetByName("Dados Para Corrigir na Matriz");

  let linhaInicial = abacpfsSemModalidade.getRange(1, 6).getValue();
  let ultimaLinha = abaMatriz.getLastRow() - (linhaInicial - 1);


  let modalidades = abaMatriz.getRange(linhaInicial, 11,ultimaLinha, 1).getValues();
  let cpfs = abaMatriz.getRange(linhaInicial, 4, ultimaLinha, 4).getValues();

  verificarExistenciaModalidade(abacpfsSemModalidade, modalidades, cpfs);
}

function verificarExistenciaModalidade(abacpfsSemModalidade, modalidades, cpfs) {
  let cpfsSemModalidade = [];

  for (let i = 0; i < modalidades.length; i++) {
    if (modalidades[i][0] || i == 443) {
      continue;
    }

    let nomeDoAlunoSemCpf = cpfs[i][0];
    let cpfDaLinha = cpfs[i][1];
    let emailDoAlunoSemCpf = cpfs[i][3];

    if (!cpfDaLinha) {
      cpfDaLinha = `CPF não cadastrado, Nome: ${nomeDoAlunoSemCpf} Email: ${emailDoAlunoSemCpf}`;
    }

    cpfsSemModalidade.push([cpfDaLinha]);

    guardarModalidadesNaAba(abacpfsSemModalidade, cpfsSemModalidade);
  }
}

function guardarModalidadesNaAba(abacpfsSemModalidade, cpfsSemModalidade) {
  if (cpfsSemModalidade.length > 0) {
    abacpfsSemModalidade.getRange(4, 3, cpfsSemModalidade.length, 1).setValues(cpfsSemModalidade);
  } else {
    Logger.log("Nenhum cpf sem modalidade preenchida");
  }
}
