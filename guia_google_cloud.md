# 🔑 Guia: Gerando sua Chave JSON do Google

Siga estes passos para que o app consiga acessar seus arquivos e planilhas com total segurança.

## Parte 1: Criar o Projeto e Ativar as APIs
1.  Acesse o [Google Cloud Console](https://console.cloud.google.com/).
2.  No topo da página, ao lado de "Google Cloud", clique em **Select a Project** > **New Project**.
3.  Dê um nome, ex: `Gerador-Jardim-Equipamentos`. Clique em **Create**.
4.  No menu lateral (esquerdo), vá em **APIs & Services** > **Library**.
5.  Pesquise por **Google Drive API** e clique em **Enable**.
6.  Volte à biblioteca, pesquise por **Google Sheets API** e clique em **Enable**.

## Parte 2: Criar a Service Account (Conta de Serviço)
1.  Vá em **APIs & Services** > **Credentials**.
2.  Clique em **+ CREATE CREDENTIALS** (no topo) e selecione **Service Account**.
3.  **Nome**: `rodrigo-app-gerador`.
4.  Clique em **CREATE AND CONTINUE**.
5.  Em "Select a role", escolha **Basic** > **Editor**. Clique em **CONTINUE** e depois em **DONE**.

## Parte 3: Baixar a Chave JSON
1.  Na lista de "Service Accounts" (no final da tela de Credentials), clique no e-mail que você acabou de criar.
2.  Vá na aba **Keys** (no topo).
3.  Clique em **ADD KEY** > **Create new key**.
4.  Certifique-se de que **JSON** está selecionado e clique em **CREATE**.
5.  O download de um arquivo `.json` será feito. Esse é o arquivo que precisamos! 🚀

---

> [!IMPORTANT]
> **COMPARTILHAMENTO:**
> Copie o e-mail da sua Service Account (ex: `rodrigo-app-gerador@...iam.gserviceaccount.com`).
> Vá na sua **Planilha** e na sua **Pasta de Modelos** no Google Drive, clique em **Compartilhar (Share)** e adicione esse e-mail como **Editor**. Só assim o app poderá ver os arquivos!
