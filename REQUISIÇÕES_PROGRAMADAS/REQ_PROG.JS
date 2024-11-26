function transferData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const programadasSheet = ss.getSheetByName("NOME DA ABA DA PLANILHA DE ORIGEM");
    const menuSheet = ss.getSheetByName("NOME DA ABA DA PLANILHA DE DESTINO");
  
    const programadasData = programadasSheet.getDataRange().getValues();
    const menuData = menuSheet.getDataRange().getValues();
  
    // Cria um conjunto para armazenar dados já existentes na aba MENU
    const menuSet = new Set(menuData.map(row => row.join(',')));
  
    // Prepara um array para os novos dados
    const newData = [];
  
    // Percorre os dados da aba PROGRAMADAS
    for (let i = 1; i < programadasData.length; i++) { // Começa de 1 para pular o cabeçalho
      const row = programadasData[i];
      const rowString = row.join(',');
  
      // Se a linha não estiver no menuSet, adiciona ao newData
      if (!menuSet.has(rowString)) {
        newData.push(row);
      }
    }
  
    // Adiciona os novos dados na aba MENU, se houver
    if (newData.length > 0) {
      menuSheet.getRange(menuData.length + 1, 1, newData.length, newData[0].length).setValues(newData);
    } else {
      Logger.log("Nenhum novo dado para transferir.");
    }
}

// ATRAVÉS DE UM FORMS PREENCHIDO COM A OPÇÃO DE REQUSIÇÃO PROGRAMADA, ESSAS INFORMAÇÕES SERÃO ADICIONANDO EM UMA PLANILHA ESPECIFICA. 
// COM O SCRIPT ACIMA, TROUXE ESSES DADOS PARA MINHA PLANILHA.