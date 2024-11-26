function atualizarColunaN() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Pega as células da coluna B da linha 4 até a última linha preenchida
    var rangeB = sheet.getRange(4, 2, sheet.getLastRow() - 3, 1); // Linha 4 até a última linha preenchida
    var valuesB = rangeB.getValues(); // Pega todos os valores da coluna B
    
    // Pega o intervalo correspondente da coluna N a partir da linha 4
    var rangeN = sheet.getRange(4, 14, sheet.getLastRow() - 3, 1); // Linha 4 até a última linha preenchida
    var valuesN = []; // Vamos armazenar os novos valores da coluna N aqui
  
    // Percorrer todas as células da coluna B e atualizar a coluna N
    for (var i = 0; i < valuesB.length; i++) {
      var valorB = valuesB[i][0].toUpperCase(); // Garantir que o valor esteja em maiúsculo
      var valorN = "";
  
      // Atribuindo o valor correto para a coluna N
      if (valorB == "ENTREGUE AO SOLICITANTE") {
        valorN = "FINALIZADO";
      } else if (valorB == "AGUARDANDO RETIRADA") {
        valorN = "RETIRAR MATERIAL NO ALMOXARIFADO";
      } else if (valorB == "EM PROCESSO DE COMPRA") {
        valorN = "AGUARDANDO CHEGADA";
      } else if (valorB == "CANCELADO") {
        valorN = "SOLICITAÇÃO CANCELADA";
      }
  
      // Adiciona o valor para a coluna N na linha correspondente
      valuesN.push([valorN]);
    }
  
    // Atualiza as células da coluna N
    rangeN.setValues(valuesN);
}

// DE ACORDO COM A INFORMAÇÃO FINAL NA COLUNA B, O STATUS FINAL DA REQUISIÇÃO NA COLUNA N SERÁ INSERIDO AUTOMÁTICAMENTE.