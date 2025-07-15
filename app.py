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

def format_pos_entrevista(df):
    """
    Formata os dados da aba "Pós entrevista".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "2 - Contexto e escopo do processo\n"
        content += f"\t2.1 - Objetivo de negócio do processo (qual valor gera?): {row.iloc[0]}\n"
        content += f"\t2.2 – Fronteiras do processo (onde começa e termina; o que fica fora): {row.iloc[1]}\n"
        content += f"\t2.3 – Stakeholders envolvidos (internos, externos, sistemas): {row.iloc[2]}\n"
        content += f"\t2.4 – Encaixe em processos 'pai' - (Indica se este BPMN será chamado como Call Activity ou é o nível mais alto.): {row.iloc[3]}\n"
        content += "\n"
    return content

def format_etapas_processo(df):
    """
    Formata os dados da aba "Etapas do Processo".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "3 - Etapas do Processo\n"
        content += f"\t3.1 - Etapa Nº: {row.iloc[1]}\n"
        content += f"\t3.2 - Nome da Etapa: {row.iloc[3]}\n"
        content += f"\t3.3 - Atividade Realizada: {row.iloc[4]}\n"
        content += f"\t3.4 - Responsavel (Pessoa/papel): {row.iloc[5]}\n"
        content += f"\t3.5 - Porque realizar essa atividade?: {row.iloc[6]}\n"
        content += f"\t3.6 - Sistema utilizado: {row.iloc[7]}\n"
        content += f"\t3.7 - Quando é executada?: {row.iloc[9]}\n"
        content += f"\t3.8 - Entrada (Documento/dados) - (Eventos disparadores típicos, documento recebido, agendamento, erro, SLA, etc.): {row.iloc[10]}\n"
        content += f"\t3.9 - Saída (Documento/dados): {row.iloc[11]}\n"
        content += f"\t3.10 - Regras de negócio aplicadas: {row.iloc[12]}\n"
        content += f"\t\t3.10.1 - Decisão automática ou humana?: {row.iloc[19]}\n"
        content += f"\t\t3.10.2 - Critérios da decisão: {row.iloc[20]}\n"
        content += f"\t\t3.10.3 - Frequência de exceções: {row.iloc[21]}\n"
        content += f"\t\t3.10.4 - Tolerância/SLA: {row.iloc[22]}\n"
        content += f"\t3.11 - Duração estimada: {row.iloc[14]}\n"
        content += f"\t3.12 - Oservações: {row.iloc[15]}\n"
        content += "\n"
    return content

def format_oportunidades_melhorias(df):
    """
    Formata os dados da aba "Oportunidades de Melhorias".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "4 - Dores e Melhorias\n"
        content += f"\t4.1 - Etapa relacionada: {row.iloc[1]}\n"
        content += f"\t4.2 - Problema ou dor atual: {row.iloc[2]}\n"
        content += f"\t4.3 - Impacto do problema: {row.iloc[3]}\n"
        content += f"\t4.4 - Sugestão de melhoria: {row.iloc[4]}\n"
        content += "\n"
    return content

def format_excecoes(df):
    """
    Formata os dados da aba "Exceções".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "5 - Exceções\n"
        content += f"\t5.1 - Etapa relacionada: {row.iloc[1]}\n"
        content += f"\t5.2 - Descricão da exceção: {row.iloc[2]}\n"
        content += f"\t5.3 - Com é tratada atualmente?: {row.iloc[3]}\n"
        content += f"\t5.4 - Quem decide?: {row.iloc[4]}\n"
        content += "\n"
    return content

def format_anexos(df):
    """
    Formata os dados da aba "Anexos".
    """
    content = ""
    for index, row in df.iloc[8:].iterrows():
        content += "6 - Anexos\n"
        content += f"\t6.1 - Nome do documento: {row.iloc[1]}\n"
        content += f"\t6.2 - Tipo (modelo, manual, formulário, etc): {row.iloc[2]}\n"
        content += f"\t6.3 - Descrição/Finalidade do uso: {row.iloc[3]}\n"
        content += f"\t6.4 - Local de armazenamento (link/pasta): {row.iloc[4]}\n"
        content += "\n"
    return content

RULES = [
    {
        "sheet_name": "Entrevista - Geral",
        "formatter": format_entrevista_geral,
    },
    {
        "sheet_name": "Pós entrevista",
        "formatter": format_pos_entrevista,
    },
    {
        "sheet_name": "Etapas do Processo",
        "formatter": format_etapas_processo,
    },
    {
        "sheet_name": "Oportunidades de Melhorias",
        "formatter": format_oportunidades_melhorias,
    },
    {
        "sheet_name": "Exceções",
        "formatter": format_excecoes,
    },
    {
        "sheet_name": "Anexos",
        "formatter": format_anexos,
    },
]


def main():
    st.title("Conversor de Excel para TXT")

    st.sidebar.title("Menu")
    if st.sidebar.button("Upload de Excel"):
        st.session_state.page = "upload"
    if st.sidebar.button("Conversão para TXT"):
        st.session_state.page = "convert"

    if "page" not in st.session_state:
        st.session_state.page = "upload"

    if st.session_state.page == "upload":
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

    elif st.session_state.page == "convert":
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
