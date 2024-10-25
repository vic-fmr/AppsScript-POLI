//Script para verificar casos de cpfs cadastrados em 2 estágios classificados com "EM ANDAMENTO" simultanemente

function verificarDuplicidade() {
    let planilha = SpreadsheetApp.getActiveSpreadsheet()
    let aba_matriz = planilha.getSheetByName("CONTROLE - 2022.2");
    let aba_cpfs_duplos = planilha.getSheetByName("Dados Para Corrigir na Matriz")
  
    let cpfs = aba_matriz.getRange(5, 5, aba_matriz.getLastRow() - 4, 1).getValues()
    let status = aba_matriz.getRange(5, 40, aba_matriz.getLastRow() - 4, 1).getValues()
  
  
  
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
  
    if (cpfs_duplos.length > 0) {
    
    let cpfs_duplos_2D = cpfs_duplos.map(cpf => [cpf]); 
    aba_cpfs_duplos.getRange(2, 1, cpfs_duplos_2D.length, 1).setValues(cpfs_duplos_2D);
  }
  }