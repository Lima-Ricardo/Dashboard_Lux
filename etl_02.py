from pandasgui import show
import pandas as pd
import hashlib

# Carregar o arquivo Excel
path = "F:\\pythonProject\\Simulação_Projeto_Interno_Tratado.xlsx"

# Carregar os dados das abas "Clientes, Empresas"
df_dados = pd.read_excel(path, sheet_name='Dados Brutos', decimal= '.')
df_clientes = pd.read_excel(path, sheet_name='ETL01', decimal= '.')
df_empresas = pd.read_excel(path, sheet_name='ETL02', decimal= '.')

#---------------------------------------------------------------------------------
#ETL do primeiro dataframe Clientes.


df1 = df_clientes.copy()
df1.columns = df1.iloc[0]
df1 = df1[1:]
df1 = df1[df1['Identificador'] != 'Colaborador Lux']

df1.reset_index(drop=True, inplace=True)


print(df1.columns)

#---------------------------------------------------------------------------------
#ETL do segundo dataframe Empresas

df2 = df_empresas.copy()
df2.columns = df2.iloc[0]
df2 = df2[1:]

#df2.reset_index(drop=True, inplace=True)

def create_unique_id(text):
    normalized_text = text.strip().lower()  # Normaliza o texto
    return hashlib.md5(normalized_text.encode()).hexdigest()

df2['ID'] = df2['Empresa'].apply(create_unique_id)

#---------------------------------------------------------------------------------
# Criar um dicionário de mapeamento de Empresa para ID
id_mapping = dict(zip(df2['Empresa'], df2['ID']))

# Adicionar a coluna 'ID' ao df1 com base na coluna 'Identificador'
df1['ID'] = df1['Identificador'].map(id_mapping)

#---------------------------------------------------------------------------------
# Sobrescrevendo os dados nas abas do arquivo Excel
with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df1.to_excel(writer, sheet_name='Clientes', index=False)
    df2.to_excel(writer, sheet_name='Empresas', index=False)
    df_dados.to_excel(writer, sheet_name='Dados Brutos', index=False)

print("Dados atualizados com sucesso.")

# Opcional: Exibir o DataFrame usando a interface gráfica do PandasGUI necessário instalação do pacote "pip install pandasgui"
gui = show(df1)