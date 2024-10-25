function abrirDialogo() {
    let htmlOutput = HtmlService.createHtmlOutputFromFile('Interface Notas')
        .setWidth(700)
        .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Insira as Notas');
  }
  
  function receberDados(nota_supervisorValor, nota_orientadorValor) {
    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let linhaAtiva = planilha.getActiveCell().getRow();
    let cpfAluno = planilha.getRange(linhaAtiva, 7).getValue(); // Obtém o CPF do aluno na linha ativa
    let linkRelatorio = planilha.getRange(linhaAtiva, 20).getValue(); // Obtém o link do relatório
  
  
    let planilhaDestino = SpreadsheetApp.openById("1-vzITec7MQhX7GYYdGaaOzFGvqYorG2lSHt2mhKSpL0").getActiveSheet();
    let ultimaLinha = planilhaDestino.getLastRow();
  
    // Lê todos os dados de uma vez só em vez de ler duas vezes
    let dadosDestino = planilhaDestino.getRange(3, 5, ultimaLinha - 2, 31).getValues(); // Lê CPFs, classificação e colunas a partir da 5ª
  
    let linhaDestino = null;
    let contadorCPF = 0; // letiável para contar quantas vezes o CPF aparece com "Obrigatório"
  
    for (let i = 0; i < dadosDestino.length; i++) {
      let cpf_destino = dadosDestino[i][0]; // Coluna do CPF
      let tipoEstagio = dadosDestino[i][6]; // Coluna 11 relativa ao tipo de estágio (coluna 6 no array zero-index)
  
      // Verifica se o CPF é igual e o estágio é "Obrigatório"
      if (cpf_destino == cpfAluno && tipoEstagio === "Obrigatório") {
        contadorCPF++; // Incrementa o contador de CPFs obrigatórios
  
        if (contadorCPF > 1) {
          // Se encontrar mais de uma linha com o mesmo CPF e estágio "Obrigatório", exibe um alerta
          SpreadsheetApp.getUi().alert('Erro: O CPF ' + cpfAluno + ' aparece mais de uma vez com estágio "Obrigatório". Verifique a planilha.');
          return; // Sai da função para interromper o processo
        }
  
        linhaDestino = i + 3; // Armazena a linha correspondente
      }
    }
  
    if (linhaDestino) {
      // Atualiza as notas e o link do relatório de uma só vez
      planilhaDestino.getRange(linhaDestino, 33, 1, 3).setValues([[linkRelatorio, nota_supervisorValor, nota_orientadorValor]]);
    }
  }
  