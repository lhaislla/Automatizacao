import pandas as pd 
import numpy as np
import plotly as pl


# Abrir o arquivo Excel
excel_file = pd.ExcelFile('./custo/custo.xlsx')
df = pd.read_excel(excel_file,header=6)
print(df.head())

# Remover as primeiras 6 linhas
df = df.drop(index=range(5))
# Remover colunas vazias
df = df.dropna(axis=1, how='all')
colunas_para_remover = ['Tp.aval.', 'Documento SD', 'Prç.padrão', 'PrçIntPer.', 'Tipo de material', 'Unnamed: 1', 'Moeda']
df = df.drop(columns=colunas_para_remover, errors='ignore')  # 'errors='ignore'' para ignorar colunas que não existem
# Remove espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()
#df.columns 
ordem_colunas = ['Material', 'Texto breve material','Estoque total', 'UMB', 'Val.total']
df = df[ordem_colunas]

# Encontrar o índice da primeira linha onde 'Val.total', 'Material' e 'Estoque total' são vazias
indice_linha_vazia = df[(df['Val.total'].isnull()) & (df['Material'].isnull()) & (df['Estoque total'].isnull())].index.min()

# Se encontrou uma linha vazia, excluir todas as linhas a partir deste índice em todas as colunas
if pd.notnull(indice_linha_vazia):
    df = df.iloc[:indice_linha_vazia, :]

caminho_xlsx = './custo/novo.xlsx'  # Salva o arquivo Excel
df.to_excel(caminho_xlsx, index=False)
