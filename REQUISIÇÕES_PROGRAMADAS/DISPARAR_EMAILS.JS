// Função personalizada que será executada ao editar a planilha
function onEditRequestStatus(e) {
    var sheet = e.source.getActiveSheet();
    var editedCell = e.range;
    
    // Verifica se a edição foi feita na coluna B (índice 2) e se o valor alterado é "AGUARDANDO RETIRADA"
    if (editedCell.getColumn() == 2 && editedCell.getValue() == "AGUARDANDO RETIRADA") {
      var row = editedCell.getRow();
      
      // Obtém os emails dos aprovadores (colunas R e S)
      var email1 = sheet.getRange(row, 18).getValue();  // Coluna R (18)
      var email2 = sheet.getRange(row, 19).getValue();  // Coluna S (19)
      
      // Obtém as informações detalhadas da requisição
      var requisicao = sheet.getRange(row, 2).getValue();  // Número da Requisição (coluna B)
      var dataCriacao = sheet.getRange(row, 4).getValue();  // Data de Criação (coluna D)
      var centroCusto = sheet.getRange(row, 4).getValue();  // Centro de Custo (coluna E)
      var nomeSolicitante = sheet.getRange(row, 6).getValue();  // Nome do Solicitante (coluna F)
      var codigoMaterial = sheet.getRange(row, 7).getValue();  // Código do Material (coluna G)
      var localRetirada = sheet.getRange(row, 8).getValue();  // Local de Retirada (coluna H)
      var descricaoMaterial = sheet.getRange(row, 9).getValue();  // Descrição do Material (coluna I)
      
      // Envia o primeiro e-mail com as informações detalhadas
      if (email1 || email2) {
        var subject = "Aprovador, a requisição está aguardando retirada";
        
        // Corpo do e-mail com emojis como ícones antes de cada tópico
        var body = "Olá,\n\nA requisição de número " + requisicao + " está aguardando retirada. Por favor, realizar a retirada ou informar o solicitante.\n\n" +
                   "Aqui estão os detalhes da requisição:\n\n" +
                   "📅 Data de Criação: " + dataCriacao + "\n" +
                   "🏷️ Centro de Custo: " + centroCusto + "\n" +
                   "👤 Nome do Solicitante: " + nomeSolicitante + "\n" +  // Nome do solicitante adicionado
                   "🔢 Código do Material: " + codigoMaterial + "\n" +
                   "📍 Local de Retirada*: " + localRetirada + "\n" +
                   "📝 Descrição do Material: " + descricaoMaterial + "\n\n" +
                   "Atenciosamente,\nAlmoxarifado Site";
        
        // Envia o e-mail para o primeiro aprovador
        if (email1) {
          MailApp.sendEmail(email1, subject, body);
        }
        
        // Envia o e-mail para o segundo aprovador, se houver
        if (email2) {
          MailApp.sendEmail(email2, subject, body);
        }
      }
      
      // Cria um gatilho de tempo para enviar lembretes a cada 3 horas
      ScriptApp.newTrigger("sendReminderEmails")
        .timeBased()
        .everyHours(3) // Envia a cada 3 horas
        .create();
    }
  
    // Verifica se a célula foi alterada para "ENTREGUE AO SOLICITANTE"
    if (editedCell.getColumn() == 2 && editedCell.getValue() == "ENTREGUE AO SOLICITANTE") {
      // Apaga todos os gatilhos de lembretes para essa linha
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == "sendReminderEmails") {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
    }
  }
  
  // Função para enviar e-mail de lembrete e agendar mais envios
  function sendReminderEmails() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();  // Obtém todos os dados da planilha
    
    // Loop através das linhas da planilha
    for (var i = 0; i < data.length; i++) {
      var status = data[i][1]; // Status na coluna B (índice 1)
      var email1 = data[i][17]; // E-mail do aprovador 1 na coluna R (índice 17)
      var email2 = data[i][18]; // E-mail do aprovador 2 na coluna S (índice 18)
      var requisicao = data[i][2]; // Número da requisição na coluna C (índice 2)
      
      // Obtém as informações detalhadas da requisição
      var dataCriacao = data[i][3];  // Data de Criação (coluna D)
      var centroCusto = data[i][4];  // Centro de Custo (coluna C)
      var nomeSolicitante = data[i][5];  // Nome do Solicitante (coluna F)
      var codigoMaterial = data[i][6];  // Código do Material (coluna G)
      var localRetirada = data[i][7];  // Local de Retirada (coluna H)
      var descricaoMaterial = data[i][8];  // Descrição do Material (coluna I)
      
      // Verifica se o status está "AGUARDANDO RETIRADA"
      if (status == "AGUARDANDO RETIRADA") {
        var subject = "Aprovador, a requisição está aguardando retirada";
        
        // Corpo do e-mail com emojis como ícones antes de cada tópico
        var body = "Olá,\n\nA requisição de número " + requisicao + " está aguardando retirada. Por favor, realizar a retirada ou informar o solicitante\n\n" +
                   "Aqui estão os detalhes da requisição:\n\n" +
                   "📅 Data de Criação: " + dataCriacao + "\n" +
                   "🏷️ Centro de Custo: " + centroCusto + "\n" +
                   "👤 Nome do Solicitante: " + nomeSolicitante + "\n" +  // Nome do solicitante adicionado
                   "🔢 Código do Material: " + codigoMaterial + "\n" +
                   "📍 Local de Retirada: " + localRetirada + "\n" +
                   "📝 Descrição do Material: " + descricaoMaterial + "\n\n" +
                   "Atenciosamente,\nAlmoxarifado";
        
        // Envia o e-mail para os aprovadores se houver e-mails válidos
        if (email1) {
          MailApp.sendEmail(email1, subject, body);
        }
        if (email2) {
          MailApp.sendEmail(email2, subject, body);
        }
      }
    }
}

// APÓS A CHEGADA DO MATERIAL COMPRADO PROGRAMADO, SEPARAMOS O MATERIAL E AGUARDAMOS O APROVADOR FAZER SUA RETIRADA.
// PARA ISSO, DESENVOLVI ESSE SCRIPT, ONDE O VALOR DA CÉLULA NA COLUNA B FOR "AGUARDANDO RETIRADA", O APROVADOR RECEBERÁ UM EMAIL AUTOMÁTICO COM AS INFORMAÇÕES DA REQUSIÇÃO.


