# Envio de Operações - Bluemetrix

Aplicativo para geração e envio automático de relatórios em PDF para assessores e consolidados.

## Como rodar localmente
1. Clone o repositório ou baixe os arquivos.
2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Crie um arquivo `.env` baseado no `.env.example` e configure com sua senha do e-mail:
   ```
   EMAIL_PASSWORD=sua_senha_aqui
   ```
4. Coloque os arquivos de dados na pasta raiz:
   - `ordens.csv`
   - `acompanhamento_de_operacoes.csv`
   - `Controle de Contratos - Atualizado 2025.xlsx`
   - `emails.csv`
5. Rode o app:
   ```bash
   streamlit run app.py
   ```

## Como fazer deploy no Streamlit Cloud
1. Crie uma conta em [https://streamlit.io/cloud](https://streamlit.io/cloud).
2. Suba os arquivos do projeto para um repositório no GitHub.
3. No painel do Streamlit Cloud, clique em **"New App"**, aponte para o seu repositório e configure:
   - **Branch**: branch principal (ex.: `main`)
   - **Main file path**: `app.py`
4. Configure as **Secrets** (variáveis de ambiente):
   - Vá em **Settings → Secrets** e adicione:
     ```
     EMAIL_PASSWORD=sua_senha_aqui
     ```
5. Clique em **Deploy**. O app ficará disponível em um link do tipo:
   ```
   https://seuapp.streamlit.app
   ```

## Estrutura do projeto
```
/envio-relatorios
├── app.py
├── assessores.py
├── ordens.csv
├── acompanhamento_de_operacoes.csv
├── Controle de Contratos - Atualizado 2025.xlsx
├── emails.csv
├── Assinatura David.jpg
├── pdfs/
├── requirements.txt
├── README.md
├── .env.example
```
