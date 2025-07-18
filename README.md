# 💰 Sistema de Aprovação de Reembolsos com Google Apps Script

Este projeto oferece uma solução robusta e automatizada para o fluxo de **aprovação de reembolsos**, utilizando a integração entre **Google Forms**, **Google Sheets** e **Google Apps Script**. Inicialmente implementado na **Medway**, o sistema visa agilizar o envio de e-mails de aprovação e o gerenciamento de status diretamente na planilha.

---

## 🚀 Visão Geral

Após um colaborador preencher um formulário de solicitação de reembolso, o sistema:

- Identifica novas submissões
- Envia e-mails personalizados para os aprovadores
- Atualiza automaticamente os status na planilha vinculada

---

## ✨ Funcionalidades

- **📧 Automação de E-mails**: Envio automático para os aprovadores com os detalhes da solicitação.
- **✅ Aprovação/Reprovação Rápida**: Botões de **"Aprovar"** e **"Reprovar"** diretamente no corpo do e-mail.
- **📊 Atualização em Tempo Real**: Planilha atualizada automaticamente com o status e a data da resposta.
- **📌 Rastreamento de E-mails**: Evita envios duplicados com marcações de e-mail enviado.
- **⚙️ Configuração Flexível**: Personalizável de acordo com as colunas e abas da sua planilha.
- **🔗 Independente de URL**: Funciona diretamente na planilha vinculada, sem depender de links fixos.

---

## ⚙️ Como Funciona

1. **📝 Submissão do Formulário**  
   Um usuário preenche o Google Formulário.

2. **⏱️ Gatilho de Tempo**  
   Um gatilho executa periodicamente a função `processarReembolsosPendentes`.

3. **🔍 Identificação de Pendências**  
   O script verifica a planilha e encontra solicitações com status "Pendente" ou vazio.

4. **📬 Envio de E-mails**  
   Um e-mail com botões HTML para "Aprovar" ou "Reprovar" é enviado ao aprovador.

5. **👆 Ação do Aprovador**  
   Ao clicar no botão, o Web App é acionado.

6. **📥 Atualização da Planilha**  
   O status e a data/hora da resposta são atualizados automaticamente.

---

## 🛠️ Configuração

### 1. Criar Formulário e Planilha

- Crie um Google Formulário e vincule-o a uma planilha (Respostas).

### 2. Editor de Apps Script

- Vá em **Extensões > Apps Script** na planilha.
- Cole o conteúdo do `Code.gs` no editor.

### 3. Ajuste as Variáveis de Configuração

```javascript
const SHEET_NAME = "Respostas ao formulário 1"; // Nome da aba
const WEB_APP_URL = ""; // Preencher após a implantação
const COLUMN_CONFIG = {
  NAME: 2,
  AMOUNT: 3,
  JUSTIFICATION: 4,
  APPROVER_EMAIL: 5,
  STATUS: 6,
  RESPONSE_DATE: 7,
  EMAIL_SENT_STATUS: 8
};
