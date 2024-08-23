import streamlit as st
import pandas as pd
import plotly.express as px

# Configuração do Streamlit para o tema escuro
st.set_page_config(page_title="Ferramenta de Gestão", layout="wide")
st.markdown(
    """
    <style>
    .css-1v3fvcr {
        background-color: #000000;
    }
    .css-1r6slb1 {
        color: #FFFFFF;
    }
    .css-1bu3u8m {
        color: #FFFFFF;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Caminho do arquivo
path = "dados.xlsx"

# Função para carregar os dados
def load_data():
    try:
        df1 = pd.read_excel(path, sheet_name='Clientes', decimal='.')
        df2 = pd.read_excel(path, sheet_name='Empresas', decimal='.')
    except FileNotFoundError:
        df1 = pd.DataFrame(columns=["Identificador", "Nome", "Cargo", "Empresa", "Telefone", "e-mail"])
        df2 = pd.DataFrame(
            columns=["Empresa", "Endereço - Rua", "Endereço - Numero", "Endereço - Estado", "Endereço - Cidade",
                     "Endereço - CEP", "Razão Social", "CNPJ", "Distribuidora", "Modalidade Tarifária", "Consumo Ponta ",
                     "Consumo Fora Ponta ", "Valor Médio da Fatura", "Gestor Responsável "])
    return df1, df2

# Função para criar gráficos com as análises
def create_charts(df1, df2):
    st.subheader('Análise Geral')

    # Cards com Análises
    col1, col2, col3 = st.columns(3)

    # Card 1: Total Consumo Ponta
    if not df2.empty and 'Consumo Ponta ' in df2.columns:
        total_consumo_ponta = df2['Consumo Ponta '].sum()
        with col1:
            st.metric(label="Total Consumo Ponta", value=f"{total_consumo_ponta: ,.2f}")

    # Card 2: Total Consumo Fora Ponta
    if not df2.empty and 'Consumo Fora Ponta ' in df2.columns:
        total_consumo_fora_ponta = df2['Consumo Fora Ponta '].sum()
        with col2:
            st.metric(label="Total Consumo Fora Ponta", value=f"{total_consumo_fora_ponta: ,.2f}")

    # Card 3: Média Valor da Fatura
    if not df2.empty and 'Valor Médio da Fatura' in df2.columns:
        media_valor_fatura = df2['Valor Médio da Fatura'].mean()
        with col3:
            st.metric(label="Média Valor da Fatura", value=f"{media_valor_fatura: ,.2f}")

    st.write("### Análise de Consumos e Valores")

    # Gráficos lado a lado
    col1, col2 = st.columns(2)

    # Gráfico de Consumo Ponta por Empresa
    if not df2.empty and 'Consumo Ponta ' in df2.columns and 'Empresa' in df2.columns:
        fig1 = px.bar(
            df2,
            x='Empresa',
            y='Consumo Ponta ',
            title='Consumo Ponta por Empresa',
            color_discrete_sequence=['rgba(255,255,255,0.8)'],
            hover_data={'Consumo Ponta ': ':.2f', 'Empresa': False}
        )
        fig1.update_layout(
            plot_bgcolor='#000000',
            paper_bgcolor='#1e1e1e',
            font_color='#FFFFFF',
            xaxis_title=None,
            yaxis_title=None
        )
        with col1:
            st.plotly_chart(fig1, use_container_width=True)

    # Gráfico de Consumo Fora Ponta por Empresa
    if not df2.empty and 'Consumo Fora Ponta ' in df2.columns and 'Empresa' in df2.columns:
        fig2 = px.bar(
            df2,
            x='Empresa',
            y='Consumo Fora Ponta ',
            title='Consumo Fora Ponta por Empresa',
            color_discrete_sequence=['rgba(255,255,255,0.8)'],
            hover_data={'Consumo Fora Ponta ': ':.2f', 'Empresa': False}
        )
        fig2.update_layout(
            plot_bgcolor='#000000',
            paper_bgcolor='#1e1e1e',
            font_color='#FFFFFF',
            xaxis_title=None,
            yaxis_title=None
        )
        with col2:
            st.plotly_chart(fig2, use_container_width=True)

    st.write("### Análise por Modalidade Tarifária")

    col1, col2 = st.columns(2)

    # Gráfico de Valor Médio da Fatura por Distribuidora
    if not df2.empty and 'Distribuidora' in df2.columns and 'Valor Médio da Fatura' in df2.columns:
        fig3 = px.bar(
            df2,
            x='Distribuidora',
            y='Valor Médio da Fatura',
            title='Valor Médio da Fatura por Distribuidora',
            color_discrete_sequence=['rgba(255,255,255,0.8)'],
            hover_data={'Valor Médio da Fatura': ':.2f', 'Distribuidora': False}
        )
        fig3.update_layout(
            plot_bgcolor='#000000',
            paper_bgcolor='#1e1e1e',
            font_color='#FFFFFF',
            xaxis_title=None,
            yaxis_title=None
        )
        with col1:
            st.plotly_chart(fig3, use_container_width=True)

    # Gráfico de Empresas por Modalidade Tarifária
    if not df2.empty and 'Modalidade Tarifária' in df2.columns:
        modalidade_counts = df2['Modalidade Tarifária'].value_counts().reset_index()
        modalidade_counts.columns = ['Modalidade Tarifária', 'Quantidade']
        fig4 = px.bar(
            modalidade_counts,
            x='Modalidade Tarifária',
            y='Quantidade',
            title='Quantidade de Empresas por Modalidade Tarifária',
            color_discrete_sequence=['rgba(255,255,255,0.8)'],
            hover_data={'Quantidade': True, 'Modalidade Tarifária': False}
        )
        fig4.update_layout(
            plot_bgcolor='#000000',
            paper_bgcolor='#1e1e1e',
            font_color='#FFFFFF',
            xaxis_title=None,
            yaxis_title=None
        )
        with col2:
            st.plotly_chart(fig4, use_container_width=True)

    st.write("### Análise de Contatos")

    # Gráfico de quantidade de contatos por empresa
    if not df1.empty and 'Empresa' in df1.columns:
        contact_counts = df1['Empresa'].value_counts().reset_index()
        contact_counts.columns = ['Empresa', 'Quantidade']
        fig5 = px.bar(
            contact_counts,
            x='Empresa',
            y='Quantidade',
            title='Quantidade de Contatos por Empresa',
            color_discrete_sequence=['rgba(255,255,255,0.8)'],
            hover_data={'Quantidade': True, 'Empresa': False}
        )
        fig5.update_layout(
            plot_bgcolor='#000000',
            paper_bgcolor='#1e1e1e',
            font_color='#FFFFFF',
            xaxis_title=None,
            yaxis_title=None
        )
        st.plotly_chart(fig5, use_container_width=True)

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

# Função para importar contatos de um arquivo
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

# Carregar dados
df1, df2 = load_data()

# Adicionar o logotipo da empresa no topo da sidebar
st.sidebar.image("LUX.webp", use_column_width=True)

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
            empresas_associadas = df2[df2['CNPJ'].isin(cnpjs)]['Empresa']
            contatos_relevantes = df1[df1['Empresa'].isin(empresas_associadas)]
            st.write(contatos_relevantes)

    st.subheader('Dashboard dos Contatos e Empresas')
    num_contatos = df1.shape[0]
    num_empresas = df2['Empresa'].nunique()
    st.write(f"**Número total de contatos:** {num_contatos}")
    st.write(f"**Número total de empresas:** {num_empresas}")

    st.write("### Lista de Contatos")
    st.write(df1)

    export_contacts(df1)

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
