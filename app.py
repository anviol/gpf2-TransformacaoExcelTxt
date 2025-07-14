import streamlit as st
import pandas as pd
import os

# Diretório para salvar os arquivos de Excel
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

def process_excel(file_path, rules):
    """
    Processa um arquivo Excel com base em um conjunto de regras.
    """
    txt_content = ""
    for rule in rules:
        try:
            df = pd.read_excel(file_path, sheet_name=rule["sheet_name"])
            txt_content += rule["formatter"](df)
        except Exception as e:
            st.error(f"Erro ao processar a aba '{rule['sheet_name']}': {e}")
    return txt_content

def format_entrevista_geral(df):
    """
    Formata os dados da aba "Entrevista - Geral".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "1 - Entrevista Geral\n"
        content += f"1.1 - Nome do entrevistado: {row.iloc[1]}\n"
        content += f"1.2 - Cargo/Função: {row.iloc[2]}\n"
        content += f"1.3 - Setor: {row.iloc[3]}\n"
        content += f"1.4 - Data da entrevista: {row.iloc[4]}\n"
        content += f"1.5 - Nome do entrevistador: {row.iloc[5]}\n"
        content += f"1.6 - Processo entrevistado: {row.iloc[6]}\n"
        content += f"1.7 - Como é medido o sucesso das suas atividades?: {row.iloc[7]}\n"
        content += "\n"
    return content

# Regras de processamento
RULES = [
    {
        "sheet_name": "Entrevista - Geral",
        "formatter": format_entrevista_geral,
    },
]

def main():
    st.title("Conversor de Excel para TXT")

    menu = ["Upload de Excel", "Conversão para TXT"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "Upload de Excel":
        st.header("Upload de Arquivo Excel")
        uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

        if uploaded_file is not None:
            file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success(f"Arquivo '{uploaded_file.name}' salvo com sucesso!")

        st.subheader("Arquivos Salvos")
        files = [f for f in os.listdir(UPLOAD_DIR) if f.endswith(".xlsx")]
        if not files:
            st.info("Nenhum arquivo Excel encontrado.")
        else:
            st.table(files)

    elif choice == "Conversão para TXT":
        st.header("Converter Excel para TXT")
        files = [f for f in os.listdir(UPLOAD_DIR) if f.endswith(".xlsx")]
        if not files:
            st.warning("Nenhum arquivo Excel para converter. Faça o upload primeiro.")
            return

        selected_file = st.selectbox("Selecione um arquivo para converter", files)
        if st.button("Converter"):
            file_path = os.path.join(UPLOAD_DIR, selected_file)
            txt_content = process_excel(file_path, RULES)

            st.text_area("Conteúdo do TXT", txt_content, height=300)

            st.download_button(
                label="Baixar TXT",
                data=txt_content,
                file_name=f"{os.path.splitext(selected_file)[0]}.txt",
                mime="text/plain",
            )

if __name__ == "__main__":
    main()
