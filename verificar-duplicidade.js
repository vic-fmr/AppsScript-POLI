//Script para verificar casos de cpfs cadastrados em 2 estágios classificados com "EM ANDAMENTO" simultanemente

function coletarDados() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet()
  let aba_matriz = planilha.getSheetByName("CONTROLE - 2022.2");
  let aba_cpfs_duplos = planilha.getSheetByName("Dados Para Corrigir na Matriz")

  let linhaInicial = aba_cpfs_duplos.getRange(1, 6).getValue();
  let ultimaLinha = aba_matriz.getLastRow() - (linhaInicial - 1);
  
  let cpfs = aba_matriz.getRange(linhaInicial, 5, ultimaLinha, 1).getValues()
  let status = aba_matriz.getRange(linhaInicial, 40, ultimaLinha, 1).getValues()

  verificarDuplicidade(aba_cpfs_duplos, cpfs, status)
}

function verificarDuplicidade(aba_cpfs_duplos, cpfs, status) {
  let cpfs_duplos = []

  for (i = 0; i < cpfs.length; i++) {

    let em_andamento = 0

    for (n = 0; n < status.length; n++) {

      if (cpfs[n][0] == cpfs[i][0] && status[n][0] == "EM ANDAMENTO") {
        em_andamento++


        if (em_andamento > 1) {
          cpfs_duplos.push(cpfs[i][0])
          break;
        }
      }
    }
  }
  guardarNaAba(aba_cpfs_duplos, cpfs_duplos)
}
function guardarNaAba(aba_cpfs_duplos, cpfs_duplos) {
  if (cpfs_duplos.length > 0) {

    let cpfs_duplos_2D = cpfs_duplos.map(cpf => [cpf]);
    aba_cpfs_duplos.getRange(4, 1, cpfs_duplos_2D.length, 1).setValues(cpfs_duplos_2D);
  }
}
