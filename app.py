import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Configuração do caminho para o arquivo
path = "C:\\Users\\rilim\\Downloads\\Simulação_Projeto_Interno_Tratado.xlsx"

# Função para carregar os dados
def load_data():
    try:
        df1 = pd.read_excel(path, sheet_name='Clientes', decimal='.')
        df2 = pd.read_excel(path, sheet_name='Empresas', decimal='.')
    except FileNotFoundError:
        df1 = pd.DataFrame(columns=["Identificador", "Nome", "Cargo", "Empresa", "Telefone", "e-mail"])
        df2 = pd.DataFrame(
            columns=["Empresa", "Endereço - Rua", "Endereço - Numero", "Endereço - Estado", "Endereço - Cidade",
                     "Endereço - CEP",
                     "Razão Social", "CNPJ", "Distribuidora", "Modalidade Tarifária", "Consumo Ponta ",
                     "Consumo Fora Ponta ",
                     "Valor Médio da Fatura", "Gestor Responsável "])
    return df1, df2

# Função para criar gráficos e análises
def create_charts(df1, df2):
    st.subheader('Análise Geral')

    # Cards com Análises
    col1, col2, col3 = st.columns(3)

    # Card 1: Total Consumo Ponta
    if not df2.empty and 'Consumo Ponta ' in df2.columns:
        total_consumo_ponta = df2['Consumo Ponta '].sum()
        with col1:
            st.metric(label="Total Consumo Ponta", value=f"{total_consumo_ponta:,.2f}")

    # Card 2: Total Consumo Fora Ponta
    if not df2.empty and 'Consumo Fora Ponta ' in df2.columns:
        total_consumo_fora_ponta = df2['Consumo Fora Ponta '].sum()
        with col2:
            st.metric(label="Total Consumo Fora Ponta", value=f"{total_consumo_fora_ponta:,.2f}")

    # Card 3: Média Valor da Fatura
    if not df2.empty and 'Valor Médio da Fatura' in df2.columns:
        media_valor_fatura = df2['Valor Médio da Fatura'].mean()
        with col3:
            st.metric(label="Média Valor da Fatura", value=f"{media_valor_fatura:,.2f}")

    st.write("### Análise de Consumos e Valores")

    # Gráficos lado a lado
    col1, col2 = st.columns(2)

    # Gráfico de Consumo Ponta por Empresa
    if not df2.empty:
        fig1, ax1 = plt.subplots(figsize=(10, 5))
        sns.barplot(x=df2['Empresa'], y=df2['Consumo Ponta '], color='#d6dbe4', ax=ax1)
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=90)
        ax1.set_title('Consumo Ponta por Empresa', color='white')
        ax1.set_facecolor('#c3e3f7')  # Cor de fundo dos gráficos
        ax1.set_frame_on(False)
        for p in ax1.patches:
            height = p.get_height()
            ax1.text(p.get_x() + p.get_width() / 2.,
                     height + 100,
                     f'{height:,.0f}',
                     ha="center", color='white', fontsize=10)
        with col1:
            st.pyplot(fig1)

    # Gráfico de Consumo Fora Ponta por Empresa
    if not df2.empty:
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        sns.barplot(x=df2['Empresa'], y=df2['Consumo Fora Ponta '], color='#9cacbd', ax=ax2)
        ax2.set_xticklabels(ax2.get_xticklabels(), rotation=90)
        ax2.set_title('Consumo Fora Ponta por Empresa', color='white')
        ax2.set_facecolor('#c3e3f7')  # Cor de fundo dos gráficos
        ax2.set_frame_on(False)
        for p in ax2.patches:
            height = p.get_height()
            ax2.text(p.get_x() + p.get_width() / 2.,
                     height + 100,
                     f'{height:,.0f}',
                     ha="center", color='white', fontsize=10)
        with col2:
            st.pyplot(fig2)

    st.write("### Análise por Modalidade Tarifária")

    col1, col2 = st.columns(2)

    # Gráfico de Valor Médio da Fatura por Distribuidora
    if not df2.empty and 'Distribuidora' in df2.columns and 'Valor Médio da Fatura' in df2.columns:
        fig3, ax3 = plt.subplots(figsize=(10, 5))
        sns.barplot(x=df2['Distribuidora'], y=df2['Valor Médio da Fatura'], color='#233c73', ax=ax3)
        ax3.set_title('Valor Médio da Fatura por Distribuidora', color='white')
        ax3.set_facecolor('#c3e3f7')  # Cor de fundo dos gráficos
        ax3.set_frame_on(False)
        for p in ax3.patches:
            height = p.get_height()
            ax3.text(p.get_x() + p.get_width() / 2.,
                     height + 500,
                     f'{height:,.0f}',
                     ha="center", color='white', fontsize=10)
        with col1:
            st.pyplot(fig3)

    # Gráfico de Empresas por Modalidade Tarifária
    if not df2.empty and 'Modalidade Tarifária' in df2.columns:
        modalidade_counts = df2['Modalidade Tarifária'].value_counts()
        fig4, ax4 = plt.subplots(figsize=(10, 5))
        sns.barplot(x=modalidade_counts.index, y=modalidade_counts.values, color='#233c73', ax=ax4)
        ax4.set_title('Quantidade de Empresas por Modalidade Tarifária', color='white')
        ax4.set_facecolor('#c3e3f7')  # Cor de fundo dos gráficos
        ax4.set_frame_on(False)
        for p in ax4.patches:
            height = p.get_height()
            ax4.text(p.get_x() + p.get_width() / 2.,
                     height + 1,
                     f'{height:,.0f}',
                     ha="center", color='white', fontsize=10)
        with col2:
            st.pyplot(fig4)

    st.write("### Análise de Contatos")

    # Gráfico de quantidade de contatos por empresa
    if not df1.empty and 'Empresa' in df1.columns:
        contact_counts = df1['Empresa'].value_counts()
        fig5, ax5 = plt.subplots(figsize=(10, 5))
        sns.barplot(x=contact_counts.index, y=contact_counts.values, color='#233c73', ax=ax5)
        ax5.set_xticklabels(ax5.get_xticklabels(), rotation=90)
        ax5.set_title('Quantidade de Contatos por Empresa', color='white')
        ax5.set_facecolor('#c3e3f7')  # Cor de fundo dos gráficos
        ax5.set_frame_on(False)
        for p in ax5.patches:
            height = p.get_height()
            ax5.text(p.get_x() + p.get_width() / 2.,
                     height + 1,
                     f'{height:,.0f}',
                     ha="center", color='white', fontsize=10)
        st.pyplot(fig5)

# Função para adicionar um contato
def add_contact(df1, nome, cargo, empresa, telefone_email, cnpj, gestor):
    novo_contato = pd.DataFrame([[None, nome, cargo, empresa, telefone_email]],
                                columns=["Identificador", "Nome", "Cargo", "Empresa", "Telefone", "e-mail"])
    df1 = pd.concat([df1, novo_contato], ignore_index=True)
    return df1

# Função para salvar os dados
def save_data(df1, df2):
    df1.to_excel(path, sheet_name='Clientes', index=False)
    df2.to_excel(path, sheet_name='Empresas', index=False)

# Função para carregar contatos de um arquivo
def upload_contacts(file):
    new_contacts = pd.read_excel(file)
    required_columns = ["Identificador", "Nome", "Cargo", "Empresa", "Telefone", "e-mail"]
    if not all(col in new_contacts.columns for col in required_columns):
        st.error(f"Arquivo deve conter as colunas: {', '.join(required_columns)}")
        return pd.DataFrame(columns=required_columns)
    return new_contacts

# Função para exportar contatos para CSV
def export_contacts(df1):
    csv = df1.to_csv(index=False)
    st.download_button(label="Exportar Contatos para CSV", data=csv, file_name='contatos.csv', mime='text/csv')

# Configurações do Streamlit
st.set_page_config(page_title="Ferramenta de Gestão", layout="wide")

# Carregar dados
df1, df2 = load_data()

# Criação da sidebar para navegação
st.sidebar.title("Navegação")
page = st.sidebar.radio("Escolha a Página", ["Análise Geral", "Contatos", "Inserção de Novo Contato"])

if page == "Análise Geral":
    create_charts(df1, df2)

elif page == "Contatos":
    st.title('Gestão de Contatos e Empresas')

    st.subheader('Gestão de Contatos')
    st.write("### Selecione um responsável")
    if 'Gestor Responsável ' in df2.columns:
        gestores = df2['Gestor Responsável '].dropna().unique()
        selected_gestor = st.selectbox('Selecionar um responsável', gestores)
        if selected_gestor:
            # Obter o CNPJ da empresa para o gestor selecionado
            cnpjs = df2[df2['Gestor Responsável '] == selected_gestor]['CNPJ']
            # Filtrar contatos pelo CNPJ das empresas associadas ao gestor
            contatos_relevantes = df1[df1['Empresa'].isin(df2[df2['CNPJ'].isin(cnpjs)]['Empresa'])]
            st.write(contatos_relevantes)

    st.subheader('Dashboard dos Contatos e Empresas')
    num_contatos = df1.shape[0]
    num_empresas = df2['Empresa'].nunique()
    st.write(f"**Número total de contatos:** {num_contatos}")
    st.write(f"**Número total de empresas:** {num_empresas}")

elif page == "Inserção de Novo Contato":
    st.title('Inserção de Novo Contato')

    # Formulário de inserção de novo contato
    with st.form(key='insert_form'):
        nome = st.text_input('Nome')
        cargo = st.text_input('Cargo')
        empresa = st.text_input('Empresa')
        telefone_email = st.text_input('Telefone e e-mail')
        cnpj = st.text_input('CNPJ')
        gestor = st.text_input('Gestor Responsável')

        submitted = st.form_submit_button('Adicionar Contato')
        if submitted:
            df1 = add_contact(df1, nome, cargo, empresa, telefone_email, cnpj, gestor)
            save_data(df1, df2)
            st.success('Contato adicionado com sucesso!')

    # Upload de planilhas
    st.write("### Upload de Planilha com Contatos")
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])
    if uploaded_file:
        new_contacts = upload_contacts(uploaded_file)
        if not new_contacts.empty:
            df1 = pd.concat([df1, new_contacts], ignore_index=True)
            save_data(df1, df2)
            st.success('Contatos carregados com sucesso!')

    # Exportar contatos para CSV
    export_contacts(df1)
