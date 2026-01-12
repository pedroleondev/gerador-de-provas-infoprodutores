# Contexto do Projeto: Gerador de Certificados e Envio de Emails

Este projeto é uma ferramenta de automação para gerar certificados personalizados em lote e enviá-los por e-mail utilizando a API do Gmail.

## 📌 Visão Geral

O sistema funciona em duas etapas principais:
1.  **Geração de Certificados:** Lê uma lista de alunos de uma planilha Excel, sobrepõe o nome de cada aluno em um modelo de certificado (imagem PNG) e salva os arquivos individualmente.
2.  **Envio de Emails:** Lê a mesma planilha para obter os emails, anexa o certificado correspondente e envia uma mensagem personalizada via Gmail.

## 🛠️ Tecnologias e Bibliotecas

*   **Python 3**
*   **Pandas & OpenPyXL:** Leitura e manipulação de arquivos Excel (`.xlsx`).
*   **Pillow (PIL):** Manipulação de imagens (escrever nomes nos certificados).
*   **Google API Client & OAuthLib:** Autenticação e envio de emails via API do Gmail.
*   **Fontes:** Utiliza fonte personalizada (`Montserrat-SemiBold.ttf`).

## 📂 Estrutura de Diretórios Importante

*   `alunos/`: Diretório de saída onde os certificados gerados (PNG) são salvos.
*   `arquivos/`: Contém a planilha de dados (ex: `WORKSHOP RUIVO EM 2H.xlsx`).
*   `certificado/`: Contém o modelo base do certificado (`certificado.png`).
*   `static/`: Possíveis arquivos estáticos (vazio ou pouco usado).
*   `log_envio_emails.txt`: Log de execução dos envios de email.
*   `token.json`: Arquivo gerado automaticamente após a autenticação no Google (NÃO COMITAR).
*   `client_secret_*.json`: Credenciais da API do Google (necessário para o primeiro acesso).

## 🚀 Scripts Principais

### 1. `gera_certificado-new.py`
**Função:** Gera os certificados.
*   **Entrada:** Planilha Excel (`arquivos\WORKSHOP RUIVO EM 2H.xlsx`, aba `APROVADAS`) e imagem base (`certificado\certificado.png`).
*   **Processamento:** Itera sobre os nomes, desenha o texto na posição `(290, 1145)` usando a fonte `Montserrat`.
*   **Saída:** Arquivos PNG na pasta `alunos/` com o nome do aluno (ex: `Fulano de Tal.png`).

### 2. `EnviarEmailGmail-Aprovados-ATUALIZACAO-22-10-2025-GPT.py`
**Função:** Envia os emails para os aprovados.
*   **Configuração:** Escopos do Gmail, caminhos de arquivo e intervalos de pausa definidos no início do script.
*   **Lógica:**
    *   Autentica no Gmail (abre navegador se necessário para gerar `token.json`).
    *   Itera sobre a planilha.
    *   Verifica se o certificado existe em `alunos/`.
    *   Envia email com anexo e corpo HTML personalizado.
    *   Registra sucesso ou erro em `log_envio_emails.txt`.

### Outros Scripts
*   `EnviarEmailGmail-Reprovados-ATUALIZACAO-22-10-2025-GPT.py`: Lógica similar para alunos reprovados (provavelmente sem anexo de certificado ou com texto diferente).
*   `gera_certificados.py`: Versão possivelmente antiga do gerador.

## ⚙️ Como Executar

1.  **Instalar Dependências:**
    ```bash
    pip install -r requirements.txt
    ```
2.  **Gerar Certificados:**
    ```bash
    python gera_certificado-new.py
    ```
3.  **Enviar Emails:**
    ```bash
    python EnviarEmailGmail-Aprovados-ATUALIZACAO-22-10-2025-GPT.py
    ```
    *   *Nota:* Na primeira execução, será necessário autenticar via navegador.

## 📝 Convenções e Notas
*   **Caminhos:** Os scripts usam caminhos relativos (ex: `arquivos\\...`), assumindo execução a partir da raiz do projeto.
*   **Logs:** Verifique sempre `log_envio_emails.txt` para debugar falhas no envio.
*   **Segurança:** Nunca comitar arquivos `token.json` ou `client_secret*.json` em repositórios públicos.
