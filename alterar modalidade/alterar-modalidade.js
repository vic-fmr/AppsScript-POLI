function abrirInterfaceModalidade() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('Interface Alteração Modalidade')
    .setWidth(800)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Preencher apenas as informações que o Termo alterar');
}

function alterarModalidade(modalidade, dataInicio, dataFinal, horario, dataFinalAnterior, novoOrientador, novoSupervisor, bolsa) {
  let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let linhaAtiva = planilha.getActiveCell().getRow();
  let cpfAluno = planilha.getRange(linhaAtiva, 7).getValue(); // Obtém o CPF do aluno na linha ativa
  let documentoPendencia = planilha.getRange(linhaAtiva, 18).getValue(); 
  let planilhaDestino = SpreadsheetApp.openById("1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0").getActiveSheet();
  let ultimaLinha = planilhaDestino.getLastRow();


  let modalidadeInversa = ""
  if (modalidade == "Obrigatório") {
    modalidadeInversa = "Não obrigatório"
  } else {
    modalidadeInversa = "Obrigatório"
  }

  let dadosDestino = planilhaDestino.getRange(3, 5, ultimaLinha - 2, 39).getValues();

  let linhaDestino = null;
  let contadorCPF = 0; 

  for (let i = 0; i < dadosDestino.length; i++) {
    let cpf_destino = dadosDestino[i][0]; 
    let tipoEstagio = dadosDestino[i][6]; 

    
    if (cpf_destino == cpfAluno && tipoEstagio === modalidadeInversa) {
      contadorCPF++; 

      if (contadorCPF > 1 && tipoEstagio === "Obrigatório") {
        SpreadsheetApp.getUi().alert('Erro: O CPF ' + cpfAluno + ' aparece mais de uma vez com estágio "Obrigatório". Verifique a planilha.');
        return;
      }

      linhaDestino = i + 3; // Armazena a linha correspondente
    }
  }

  if (linhaDestino) {
    let linhaInteiraRange = planilhaDestino.getRange(linhaDestino, 1, 1, planilhaDestino.getLastColumn());
    let linhaInteira = linhaInteiraRange.getValues();
    let formulasStatusEstagio = planilhaDestino.getRange(linhaDestino, 40, 1 , 2).getFormulasR1C1();
    let linhaInserida = linhaDestino + 1


    

      let bolsaRange = planilhaDestino.getRange(linhaInserida, 16);
      let horarioRange = planilhaDestino.getRange(linhaInserida, 18);
      let novoSupervisorRange = planilhaDestino.getRange(linhaInserida, 14)
      let linhaInseridaRange = planilhaDestino.getRange(linhaInserida, 1, 1, planilhaDestino.getLastColumn())
      let modalidadeRange = planilhaDestino.getRange(linhaInserida, 11)
      let dataInicioRange = planilhaDestino.getRange(linhaInserida, 12)
      let dataFinalRange = planilhaDestino.getRange(linhaInserida, 13)
      let novoOrientadorRange = planilhaDestino.getRange(linhaInserida, 15)
      let documentoPendenciaRange = planilhaDestino.getRange(linhaInserida, 32);
      let formulasStatusEstagioRange = planilhaDestino.getRange(linhaInserida, 40, 1, 2)

      planilhaDestino.insertRowAfter(linhaDestino)
      linhaInseridaRange.setValues(linhaInteira);
      formulasStatusEstagioRange.setFormulasR1C1(formulasStatusEstagio)
      
      modalidadeRange.setValue(modalidade)

      if (novoOrientador) {
        novoOrientadorRange.setValue(novoOrientador)
      }
      if (novoSupervisor) {
        novoSupervisorRange.setValue(novoSupervisor)
      }
      if (dataInicio) {
        dataInicioRange.setValue(dataInicio)
      }
      if (dataFinal) {
        dataFinalRange.setValue(dataFinal)
      }
      if (horario) {
        horarioRange.setValue(horario)
      }

      planilhaDestino.getRange(linhaDestino, 13).setValue(dataFinalAnterior);
      documentoPendenciaRange.setValue(documentoPendencia)

      if (bolsa) {
        bolsaRange.setValue(bolsa)
      }


  } else{
    SpreadsheetApp.getUi().alert('Aluno Não Encontrado');
  }
}
