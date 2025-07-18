/**
 * @OnlyCurrentDoc
 *
 * Este script automatiza um fluxo de aprovação de reembolsos utilizando Google Forms,
 * Sheets e Apps Script.
 *
 * FUNCIONAMENTO:
 * 1. Um gatilho de tempo aciona a função processarReembolsosPendentes em intervalos regulares.
 * 2. A função processarReembolsosPendentes lê a planilha, identifica novas submissões (status Pendente)
 * E que ainda não tiveram o e-mail de aprovação enviado.
 * 3. Para cada submissão pendente e não enviada, ela coleta os dados, formata um e-mail e o envia para o aprovador indicado.
 * 4. O e-mail contém dois botões: "Aprovar" e "Reprovar". Cada botão possui um link para este script
 * publicado como Web App, passando a ação (aprovar/reprovar) e o número da linha na planilha.
 * 5. Ao clicar em um dos botões, a função doGet é executada.
 * 6. A função doGet atualiza a planilha de respostas com o status ("Aprovado" ou "Reprovado") e a
 * data/hora da ação, e exibe uma mensagem de confirmação no navegador.
 *
 * ===================================================================================
 * GUIA DE CONFIGURAÇÃO PARA USO (PARA REPOSITÓRIO GITHUB)
 * ===================================================================================
 *
 * 1. Crie um Formulário Google e vincule-o a uma Planilha Google.
 * 2. No editor da Planilha Google, vá em "Extensões" > "Apps Script".
 * 3. Copie e cole este código no editor do Apps Script, substituindo qualquer código existente.
 * 4. Ajuste as variáveis na seção "ÁREA DE CONFIGURAÇÃO" abaixo conforme sua necessidade.
 * - SHEET_NAME: O nome exato da aba da sua planilha onde as respostas são salvas.
 * - WEB_APP_URL: Esta URL será gerada APÓS a implantação do script.
 * - COLUMN_CONFIG: Mapeie os números das colunas e seus cabeçalhos para corresponderem
 * às colunas do seu formulário e às colunas gerenciadas pelo script (Status, Data da Resposta, Status do Email).
 *
 * 5. IMPLANTAÇÃO DO WEB APP:
 * - No editor do Apps Script, clique em "Implantar" > "Nova implantação".
 * - Selecione o tipo de implantação como "App da Web".
 * - Configure o acesso:
 * - "Executar como": Sua conta (o e-mail que está executando o script).
 * - "Quem tem acesso": "Qualquer pessoa" ou "Qualquer pessoa, inclusive anônimos" (necessário para os links funcionarem via e-mail).
 * - Clique em "Implantar".
 * - Uma URL será gerada. COPIE ESSA URL e cole-a na constante `WEB_APP_URL` abaixo.
 * - Clique em "Concluído".
 *
 * 6. CONFIGURAÇÃO DO GATILHO DE TEMPO:
 * - No editor do Apps Script, no menu lateral esquerdo, clique no ícone de "Relógio" (Gatilhos).
 * - Clique em "Adicionar Gatilho" (canto inferior direito).
 * - Configure:
 * - "Escolha qual função será executada": `processarReembolsosPendentes`
 * - "Escolha qual tipo de implantação deve ser executado": `Head`
 * - "Selecione o tipo de evento": `Baseado no tempo`
 * - "Selecione o tipo de gatilho baseado no tempo": `Cronômetro de horas` (ou o intervalo desejado, ex: `A cada 5 minutos`, `A cada hora`).
 * - Clique em "Salvar". Você precisará autorizar o script na primeira vez.
 *
 * Agora o script está configurado para enviar e-mails de aprovação automaticamente!
 */

// ===================================================================================
// ÁREA DE CONFIGURAÇÃO - PREENCHA COM SUAS INFORMAÇÕES
// ===================================================================================

// Coloque o nome exato da aba (página) da sua planilha onde as respostas são salvas.
const SHEET_NAME = "Respostas ao formulário 1"; // Exemplo: "Respostas ao formulário 1"

// Coloque a URL do seu Web App, gerada após a publicação do script.
// INSTRUÇÕES: No editor, clique em "Implantar" > "Nova implantação" > Selecione "App da Web"
// > Configure o acesso e clique em "Implantar". Copie a URL gerada e cole aqui.
// Exemplo: "https://script.google.com/macros/s/SUA_ID_DE_IMPLANTACAO/exec"
const WEB_APP_URL = ""; // DEVE SER PREENCHIDA APÓS A IMPLANTAÇÃO!


// Hash para configurar as colunas na planilha.
// Isso centraliza a configuração e facilita a manutenção.
// AJUSTE OS 'columnNumber' PARA CORRESPONDEREM ÀS SUAS COLUNAS NA PLANILHA.
// A contagem de colunas começa em 1 (A=1, B=2, etc.).
const COLUMN_CONFIG = {
  // Coluna do Carimbo de Data/Hora (geralmente a primeira coluna de um formulário)
  TIMESTAMP: {
    columnNumber: 1, // Coluna A
    header: "Carimbo de data/hora"
  },
  NOME_SOLICITANTE: {
    columnNumber: 11, // Exemplo: Coluna K - Ajuste conforme seu formulário
    header: "Nome completo:"
  },
  TIPO_DESPESA: {
    columnNumber: 4, // Exemplo: Coluna D - Ajuste conforme seu formulário
    header: "Tipo de despesa:"
  },
  DEMANDA: {
    columnNumber: 12, // Exemplo: Coluna L - Ajuste conforme seu formulário
    header: "Seu gasto é referente a qual demanda?"
  },
  DESCRICAO: {
    columnNumber: 5, // Exemplo: Coluna E - Ajuste conforme seu formulário
    header: "Descrição detalhada da despesa: \n(em casos de quilometragem, inserir km rodado para cálculo do reembolso neste campo)"
  },
  VALOR: {
    columnNumber: 9, // Exemplo: Coluna I - Ajuste conforme seu formulário
    header: "Valor total da(s) despesa(s) (em R$):"
  },
  EMAIL_APROVADOR: {
    columnNumber: 7, // Coluna G - E-mail do aprovador.
    header: "E-mail do aprovador:"
  },
  COMPROVANTE_URL: {
    columnNumber: 8, // Exemplo: Coluna H - Ajuste conforme seu formulário
    header: "Comprovante de pagamento\n(nota fiscal, recibo, etc.):"
  },
  STATUS: {
    columnNumber: 14, // Coluna N (Essa coluna será gerenciada pelo script)
    header: "Status da Aprovação"
  },
  RESPONSE_DATE: {
    columnNumber: 15, // Coluna O (Essa coluna será gerenciada pelo script)
    header: "Data da Resposta"
  },
  EMAIL_SENT_STATUS: { // NOVO: Coluna para rastrear se o e-mail foi enviado
    columnNumber: 16, // Coluna P (Ajuste conforme sua planilha)
    header: "Status do Email de Aprovação"
  }
};

// ===================================================================================
// FUNÇÃO PRINCIPAL - ACIONADA PELO GATILHO DE TEMPO
// ===================================================================================

/**
 * Função executada automaticamente por um gatilho de tempo para processar reembolsos pendentes.
 */
function processarReembolsosPendentes() {
  try {
    Logger.log("▶ Iniciando processamento de reembolsos pendentes.");

    // O script opera na planilha ativa à qual está vinculado.
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log(`❌ Erro: A aba com o nome "${SHEET_NAME}" não foi encontrada na planilha. Verifique a configuração de SHEET_NAME.`);
      return;
    }

    const lastRow = sheet.getLastRow();

    // Define a linha de início para leitura dos dados (346, como solicitado)
    const START_ROW = 346;

    // Se a última linha da planilha for menor que a linha de início, não há dados para processar
    if (lastRow < START_ROW) {
      Logger.log(`Nenhuma solicitação encontrada a partir da linha ${START_ROW} na planilha.`);
      return;
    }

    // Calcula o número de linhas a serem lidas
    const numRowsToRead = lastRow - START_ROW + 1;
    // Se não houver linhas válidas para ler, retorne
    if (numRowsToRead <= 0) {
      Logger.log(`Nenhuma solicitação encontrada a partir da linha ${START_ROW} para processar.`);
      return;
    }

    // Lê os dados da planilha a partir da START_ROW até a lastRow, da coluna 1 até a última coluna necessária.
    const maxColumnNeeded = Math.max(
      COLUMN_CONFIG.TIMESTAMP.columnNumber,
      COLUMN_CONFIG.NOME_SOLICITANTE.columnNumber,
      COLUMN_CONFIG.TIPO_DESPESA.columnNumber,
      COLUMN_CONFIG.DEMANDA.columnNumber,
      COLUMN_CONFIG.DESCRICAO.columnNumber,
      COLUMN_CONFIG.VALOR.columnNumber,
      COLUMN_CONFIG.EMAIL_APROVADOR.columnNumber,
      COLUMN_CONFIG.COMPROVANTE_URL.columnNumber,
      COLUMN_CONFIG.STATUS.columnNumber,
      COLUMN_CONFIG.RESPONSE_DATE.columnNumber,
      COLUMN_CONFIG.EMAIL_SENT_STATUS.columnNumber // Inclui a nova coluna
    );

    const dataRange = sheet.getRange(START_ROW, 1, numRowsToRead, maxColumnNeeded);
    const data = dataRange.getValues();

    let emailsSentCount = 0;
    let processedRowsCount = 0;

    // Itera sobre cada linha de dados lida
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // currentRowNumber é o número real da linha na planilha
      const currentRowNumber = START_ROW + i;
      processedRowsCount++;

      // Pega o status atual da linha e o status de e-mail enviado usando os índices do array `row`
      const currentStatus = row[COLUMN_CONFIG.STATUS.columnNumber - 1];
      const emailSentStatus = row[COLUMN_CONFIG.EMAIL_SENT_STATUS.columnNumber - 1];

      // Processa a linha apenas se o status estiver vazio (nova submissão) ou "Pendente"
      // E se o e-mail de aprovação AINDA NÃO FOI ENVIADO para esta solicitação (emailSentStatus vazio)
      if ((currentStatus === "" || currentStatus === "⏳ Pendente") && emailSentStatus === "") {
        Logger.log(`▶ Processando nova solicitação na linha ${currentRowNumber} (Status atual: '${currentStatus}', Status Email Enviado: '${emailSentStatus}')`);

        // Mapeia os dados da linha para variáveis usando as colunas da hash
        const nomeSolicitante = row[COLUMN_CONFIG.NOME_SOLICITANTE.columnNumber - 1] || "Não informado";
        const tipoDespesa = row[COLUMN_CONFIG.TIPO_DESPESA.columnNumber - 1] || "Não informado";
        const demanda = row[COLUMN_CONFIG.DEMANDA.columnNumber - 1] || "Não informada";
        const descricao = row[COLUMN_CONFIG.DESCRICAO.columnNumber - 1] || "Não informada";
        const valor = row[COLUMN_CONFIG.VALOR.columnNumber - 1] || "0";
        const emailAprovador = row[COLUMN_CONFIG.EMAIL_APROVADOR.columnNumber - 1];
        const comprovanteUrl = row[COLUMN_CONFIG.COMPROVANTE_URL.columnNumber - 1] || "";

        // Validação essencial: se não houver e-mail do aprovador, o fluxo não pode continuar para esta linha.
        if (!emailAprovador) {
          Logger.log(`❌ E-mail do aprovador não fornecido para a linha ${currentRowNumber}. Marcando como erro.`);
          // Marca o status de aprovação como erro e o status de e-mail enviado para evitar re-processamento
          sheet.getRange(currentRowNumber, COLUMN_CONFIG.STATUS.columnNumber).setValue("⚠️ Erro: Sem Aprovador");
          sheet.getRange(currentRowNumber, COLUMN_CONFIG.EMAIL_SENT_STATUS.columnNumber).setValue("Erro - Sem Aprovador");
          continue; // Pula para a próxima linha
        }

        // Define o status inicial na planilha como "Pendente" (se ainda não estiver)
        // E marca o status de e-mail enviado com a data/hora atual
        sheet.getRange(currentRowNumber, COLUMN_CONFIG.STATUS.columnNumber).setValue("⏳ Pendente");
        sheet.getRange(currentRowNumber, COLUMN_CONFIG.EMAIL_SENT_STATUS.columnNumber).setValue("Enviado em " + new Date().toLocaleString()); // Marca e-mail como enviado

        // Cria o corpo do e-mail em HTML.
        const htmlBody = createEmailHtml(currentRowNumber, nomeSolicitante, tipoDespesa, demanda, descricao, valor, comprovanteUrl);

        // Envia o e-mail para o aprovador.
        MailApp.sendEmail({
          to: emailAprovador,
          subject: `Solicitação de Aprovação de Reembolso – ${nomeSolicitante}`,
          htmlBody: htmlBody,
        });

        Logger.log("✅ E-mail de aprovação enviado para " + emailAprovador + " para a linha " + currentRowNumber);
        emailsSentCount++;
      } else {
        Logger.log(`▶ Linha ${currentRowNumber} ignorada (Status: '${currentStatus}', Email Enviado: '${emailSentStatus}'). Já processada ou e-mail já enviado.`);
      }
    }
    Logger.log(`▶ Processamento concluído. ${emailsSentCount} novos e-mails de reembolso foram enviados. Total de ${processedRowsCount} linhas verificadas.`);

  } catch (error) {
    Logger.log("❌ Erro na função processarReembolsosPendentes: " + error.toString());
  }
}

// ===================================================================================
// WEB APP - PROCESSA AS AÇÕES DE APROVAÇÃO E REPROVAÇÃO
// ===================================================================================

/**
 * Função executada quando o Web App é acessado via GET (clique nos links do e-mail).
 * @param {Object} e O objeto de evento passado pelo serviço do Web App.
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const row = parseInt(e.parameter.row, 10);
    let message = "Ação inválida ou não reconhecida.";
    let status = "";
    let title = "Erro";

    // Valida se os parâmetros necessários foram recebidos.
    if (!action || !row) {
      return HtmlService.createHtmlOutput("<h1>Erro</h1><p>Parâmetros 'action' e 'row' são necessários.</p>");
    }

    // O script opera na planilha ativa à qual está vinculado.
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log(`❌ Erro: A aba com o nome "${SHEET_NAME}" não foi encontrada na planilha durante doGet.`);
      return createResponsePage("Erro Crítico", `A aba "${SHEET_NAME}" não foi encontrada. Verifique a configuração.`);
    }

    // Usa as colunas da configuração da hash
    const statusColumn = COLUMN_CONFIG.STATUS.columnNumber;
    const responseDateColumn = COLUMN_CONFIG.RESPONSE_DATE.columnNumber;

    // Verifica se a solicitação já foi processada para evitar duplicação.
    const currentStatus = sheet.getRange(row, statusColumn).getValue();
    if (currentStatus !== "⏳ Pendente") {
      title = "Atenção";
      message = "Esta solicitação já foi respondida anteriormente. Nenhuma nova ação foi tomada.";
      return createResponsePage(title, message);
    }

    // Processa a ação (aprovar ou reprovar).
    if (action === "approve") {
      status = "✅ Aprovado";
      title = "Sucesso!";
      message = "A solicitação de reembolso foi <strong>APROVADA</strong> com sucesso.";
    } else if (action === "reject") {
      status = "❌ Reprovado";
      title = "Sucesso!";
      message = "A solicitação de reembolso foi <strong>REPROVADA</strong> com sucesso.";
    }

    // Se uma ação válida foi encontrada, atualiza a planilha.
    if (status) {
      sheet.getRange(row, statusColumn).setValue(status);
      sheet.getRange(row, responseDateColumn).setValue(new Date());
      Logger.log(`▶ Linha ${row} atualizada para "${status}"`);
    }

    return createResponsePage(title, message);

  } catch (error) {
    Logger.log("❌ Erro na função doGet: " + error.toString());
    return createResponsePage("Erro Crítico", "Ocorreu um erro ao processar sua solicitação: " + error.toString());
  }
}


// ===================================================================================
// FUNÇÕES AUXILIARES
// ===================================================================================

/**
 * Cria o conteúdo HTML para o e-mail de aprovação.
 * @param {number} row O número da linha na planilha.
 * @param {string} nome Nome do solicitante.
 * @param {string} tipo Tipo de despesa.
 * @param {string} demanda Demanda relacionada à despesa.
 * @param {string} descricao Descrição da despesa.
 * @param {string} valor Valor da despesa.
 * @param {string} comprovante URL do comprovante.
 * @returns {string} O conteúdo HTML do e-mail.
 */
function createEmailHtml(row, nome, tipo, demanda, descricao, valor, comprovante) {
  // Verifica se WEB_APP_URL foi configurada
  if (!WEB_APP_URL) {
    Logger.log("❌ Erro: WEB_APP_URL não está configurada. Os links de aprovação/reprovação não funcionarão.");
    // Retorna um corpo de e-mail simplificado ou com aviso
    return `
      <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
        <p>Olá,</p>
        <p>Uma nova solicitação de reembolso foi submetida e requer sua aprovação.</p>
        <p style="color: red;"><strong>AVISO: A URL do Web App não foi configurada. Os botões de aprovação/reprovação não funcionarão.</strong></p>
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
          <tr style="background-color: #f2f2f2;">
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Solicitante:</td>
            <td style="padding: 10px; border: 1px solid #ddd;">${nome}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Tipo de Despesa:</td>
            <td style="padding: 10px; border: 1px solid #ddd;">${tipo}</td>
          </tr>
          <tr style="background-color: #f2f2f2;">
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Demanda:</td>
            <td style="padding: 10px; border: 1px solid #ddd;">${demanda}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Descrição:</td>
            <td style="padding: 10px; border: 1px solid #ddd;">${descricao}</td>
          </tr>
          <tr style="background-color: #f2f2f2;">
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Valor:</td>
            <td style="padding: 10px; border: 1px solid #ddd;">R$ ${valor}</td>
          </tr>
          <tr>
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Comprovante:</td>
            <td style="padding: 10px; border: 1px solid #ddd;"><a href="${comprovante}" target="_blank">Visualizar Comprovante</a></td>
          </tr>
        </table>
        <p>Por favor, acesse a planilha para aprovar ou reprovar manualmente.</p>
      </div>
    `;
  }

  const approveUrl = `${WEB_APP_URL}?action=approve&row=${row}`;
  const rejectUrl = `${WEB_APP_URL}?action=reject&row=${row}`;

  return `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
      <p>Olá,</p>
      <p>Uma nova solicitação de reembolso foi submetida e requer sua aprovação.</p>
      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Solicitante:</td>
          <td style="padding: 10px; border: 1px solid #ddd;">${nome}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Tipo de Despesa:</td>
          <td style="padding: 10px; border: 1px solid #ddd;">${tipo}</td>
        </tr>
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Demanda:</td>
          <td style="padding: 10px; border: 1px solid #ddd;">${demanda}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Descrição:</td>
          <td style="padding: 10px; border: 1px solid #ddd;">${descricao}</td>
        </tr>
        <tr style="background-color: #f2f2f2;">
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Valor:</td>
          <td style="padding: 10px; border: 1px solid #ddd;">R$ ${valor}</td>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold;">Comprovante:</td>
          <td style="padding: 10px; border: 1px solid #ddd;"><a href="${comprovante}" target="_blank">Visualizar Comprovante</a></td>
        </tr>
      </table>
      <p style="text-align: center; margin-top: 25px;">
        <a href="${approveUrl}" style="background-color: #4CAF50; color: white; padding: 12px 25px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-size: 16px; margin-right: 10px;">✅ Aprovar</a>
        <a href="${rejectUrl}" style="background-color: #f44336; color: white; padding: 12px 25px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-size: 16px;">❌ Reprovar</a>
      </p>
    </div>
  `;
}

/**
 * Cria uma página HTML de resposta para o usuário após o clique.
 * @param {string} title O título da página.
 * @param {string} message A mensagem a ser exibida na página.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} O objeto HTML para exibição.
 */
function createResponsePage(title, message) {
  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; background-color: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
          .container { background-color: white; padding: 40px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); text-align: center; max-width: 500px; }
          h1 { color: #333; }
          p { color: #555; font-size: 1.1em; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>${title}</h1>
          <p>${message}</p>
        </div>
      </body>
    </html>
  `);
}
