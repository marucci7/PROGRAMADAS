function copiarValoress() {
    try {
      var ssDestino = SpreadsheetApp.getActiveSpreadsheet();
      var idOrigem = ""; // ID da planilha de origem
      var origem = SpreadsheetApp.openById(idOrigem).getSheetByName("ORIGEM"); // Aba de origem
      var destino = ssDestino.getSheetByName("DESTINO"); // Aba de destino
      
      // Verifica se as abas foram encontradas
      if (!origem || !destino) {
        Logger.log("Aba de origem ou destino não encontrada.");
        return;
      }
  
      var intervaloOrigem = origem.getDataRange(); // Pega todos os dados da origem
      var valoresOrigem = intervaloOrigem.getValues(); // Obtém os valores da origem
  
      // Número de linhas da origem e da planilha de destino
      var numLinhasOrigem = valoresOrigem.length;
      var ultimaLinhaDestino = destino.getLastRow(); // Última linha com dados na planilha de destino
      Logger.log("Última linha na planilha de destino: " + ultimaLinhaDestino);
  
      // Se a planilha de destino estiver vazia, começa da primeira linha
      if (ultimaLinhaDestino == 0) {
        ultimaLinhaDestino = 1;
      } else {
        ultimaLinhaDestino += 1; // Coloca na linha seguinte à última já preenchida
      }
  
      // Copiar apenas as novas linhas (aquelas que ainda não foram copiadas)
      var linhasNovas = [];
      for (var i = ultimaLinhaDestino - 1; i < numLinhasOrigem; i++) {
        linhasNovas.push(valoresOrigem[i]);
      }
  
      // Dividir o carregamento de dados em blocos menores (por exemplo, 20 linhas de cada vez)
      var tamanhoLote = 20; // Tamanho do lote para evitar ultrapassar os limites
      for (var i = 0; i < linhasNovas.length; i += tamanhoLote) {
        var lote = linhasNovas.slice(i, i + tamanhoLote); // Cria um lote de linhas
        destino.getRange(ultimaLinhaDestino, 1, lote.length, lote[0].length).setValues(lote);
        ultimaLinhaDestino += lote.length; // Atualiza a posição para a próxima linha
      }
  
      Logger.log("Novos dados copiados com sucesso.");
      
      // Agora, verificamos se alguma linha que já foi copiada foi alterada na origem
      for (var j = 0; j < ultimaLinhaDestino - 1; j++) {
        var linhaOrigem = valoresOrigem[j];
        var linhaDestino = destino.getRange(j + 1, 1, 1, linhaOrigem.length).getValues()[0];
  
        if (JSON.stringify(linhaOrigem) !== JSON.stringify(linhaDestino)) {
          destino.getRange(j + 1, 1, 1, linhaOrigem.length).setValues([linhaOrigem]);
          Logger.log("Linha " + (j + 1) + " atualizada.");
        }
      }
  
    } catch (e) {
      Logger.log("Erro: " + e.message);
    }
}

// EM OUTRO FORMS DO GOOGLE, O SUPERVISOR E GERENTE QUE SÃO RESPONSÁVEIS PELA APROVAÇÃO DA REQUISIÇÃO, OS DADOS SÃO INSERIDOS EM OUTRA ABA.
// TROUXE TAMBÉM ESSES DADOS PARA OUTRA ABA DA MINHA PLANILHA INTERLIGANDO O NÚMERO DA REQUISIÇÃO PARA ATIVAR O PROCV.