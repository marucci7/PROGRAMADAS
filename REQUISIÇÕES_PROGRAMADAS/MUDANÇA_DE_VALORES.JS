function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var value = range.getValue();
    
    // Define os valores que causam alteração na coluna B (2ª)
    var valoresCancelados = [
      "SOLICITAÇÃO CANCELADA AUTOMATICAMENTE",
      "SOLICITAÇÃO REPROVADA PELO APROVADOR",
      "REQUISIÇÃO REPROVADA, CÓDIGO INCORRETO",
      "REPROVADO"
    ];
    
    // Verifica se a edição ocorreu na coluna P (16ª)
    if (range.getColumn() === 16) { // Coluna P
      var row = range.getRow();
      
      // Verifica se o valor editado corresponde a algum dos valores definidos
      if (valoresCancelados.indexOf(value) !== -1) {
        // Altera a coluna B da linha editada para "CANCELADO"
        sheet.getRange(row, 2).setValue("CANCELADO");
      }
    }
  
    // Verifica todas as células da coluna P e altera a coluna B
    var lastRow = sheet.getLastRow(); // Última linha com dados na planilha
    for (var i = 1; i <= lastRow; i++) {
      var cellValue = sheet.getRange(i, 16).getValue(); // Valor da célula na coluna P (16ª)
      
      // Se o valor da célula for um dos valores cancelados, altera a coluna B
      if (valoresCancelados.indexOf(cellValue) !== -1) {
        sheet.getRange(i, 2).setValue("CANCELADO"); // Coluna B = "CANCELADO"
      }
    }
}

// ESSA FUNÇÃO PERMITE ATUALIZAR A PLANILHA AUTOMÁTICAMENTE, QUANDO O VALOR DAS CÉLULAS DA COLUNA P FOR CONFORME OS VALORESCANCELADOS.
// A COLUNA B MUDA O VALOR DA CÉLULA PARA "CANCELADO".
// COM A FORMATAÇÃO CONDICIONAL, INSERI O EFEITO ONDE A LINHA MUDA PARA A COR VERMELHA, FICANDO VISIVELMENTE O CANCELAMENTO DA REQUISIÇÃO.