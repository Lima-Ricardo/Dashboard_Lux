from pandasgui import show
import pandas as pd

# Caminho do arquivo Excel
path = "F:\\pythonProject\\Simulação_Projeto_Interno.xlsx"

#---------------------------------------------------------------------------------
# Carregando o dataset da aba "Dados"
df_dados = pd.read_excel(path, sheet_name='Dados', decimal='.')

#---------------------------------------------------------------------------------
#Separando os dados para a criação da aba 01
colunas_aba01 = ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4',
                 'Unnamed: 5', 'Unnamed: 6']
df_aba01 = df_dados[colunas_aba01]

#---------------------------------------------------------------------------------
#Separando os dados para a criação da aba 02
colunas_aba02 = ['Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11',
                 'Unnamed: 12', 'Unnamed: 13', 'Unnamed: 14', 'Unnamed: 15',
                 'Unnamed: 16', 'Unnamed: 17', 'Unnamed: 18', 'Unnamed: 19',
                 'Unnamed: 20', 'Unnamed: 21']
df_aba02 = df_dados[colunas_aba02]

#---------------------------------------------------------------------------------
# Criando novas abas no excel para posteriormente aplicar o etl nos dados
with pd.ExcelWriter('Simulação_Projeto_Interno_Tratado.xlsx') as writer:
    df_aba01.to_excel(writer, sheet_name='ETL01', index=False)
    df_aba02.to_excel(writer, sheet_name='ETL02', index=False)
    df_dados.to_excel(writer, sheet_name='Dados Brutos', index=False)

print(df_dados.columns)

# Mostrar os dados com pandasgui
gui = show(df_dados)
