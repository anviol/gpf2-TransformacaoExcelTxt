# Conversor de Excel para TXT

Esta é uma aplicação Streamlit que permite fazer o upload de arquivos Excel, processá-los com base em regras predefinidas e convertê-los para o formato TXT.

## Como executar a aplicação

1.  **Instale as dependências:**

    ```bash
    pip install streamlit pandas openpyxl
    ```

2.  **Execute a aplicação:**

    ```bash
    streamlit run app.py
    ```

A aplicação estará disponível em `http://localhost:8501`.

## Funcionalidades

*   **Upload de Excel:** Faça o upload de arquivos no formato `.xlsx`.
*   **Listagem de Arquivos:** Visualize todos os arquivos que foram enviados.
*   **Conversão para TXT:** Selecione um arquivo e converta-o para TXT com base nas regras definidas no código.
*   **Download de TXT:** Baixe o arquivo TXT gerado.