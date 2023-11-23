import pandas as pd 
from datetime import date
from openpyxl import load_workbook, Workbook
import os
#Script de tratamento do arquivo  de Custo.xls

# Abre o arquivo Excel
excel_file = pd.ExcelFile('./custo.xlsx')
df = pd.read_excel(excel_file,header=6)
print(df.head())

# Remove as primeiras 6 linhas
df = df.drop(index=range(5))
# Remove colunas vazias
df = df.dropna(axis=1, how='all')
colunas_para_remover = ['Tp.aval.', 'Documento SD', 'Prç.padrão', 'PrçIntPer.', 'Tipo de material', 'Unnamed: 1', 'Moeda']
df = df.drop(columns=colunas_para_remover, errors='ignore')  # 'errors='ignore'' para ignorar colunas que não existem

# Remove espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()
#df.columns 
ordem_colunas = ['Material', 'Texto breve material','Estoque total', 'UMB', 'Val.total']
df = df[ordem_colunas]

# Encontra o índice da primeira linha onde 'Material' é vazia e a próxima linha da coluna 'Material' também é vazia
indice_linha_vazia = df['Material'][df['Material'].isnull() & df['Material'].shift(-1).isnull()].index.min()
# Se encontrou uma linha vazia, exclui todos os conteúdos das colunas após este índice
if pd.notnull(indice_linha_vazia):
    #colunas_para_manter = ['Material', 'OutrasColunasNecessarias']  # Adiciona as outras colunas que  quer manter
    df = df.loc[:indice_linha_vazia]  # Mantém as linhas até o índice da linha vazia (incluindo a linha vazia)

#Lê a base onde tem as classes 
excel_file_classe= pd.ExcelFile('.\classe.xlsx')
df_classe = pd.read_excel(excel_file_classe,header=0)

#Relaciona os materiais aos Tipos e Classes 
df = pd.merge(df, df_classe[['Material', 'Tipo', 'Classe']], on='Material', how='left')
#Lê o  arquivo base com oq ue é desconsiderado no relatório
excel_file_remover= pd.ExcelFile('./remover.xlsx')
#Remove as sandalias e Kit Dupé
df_remover= pd.read_excel(excel_file_remover,header=0)
df_remover.columns

# Usa o merge para encontrar as linhas que estão apenas em df e não em df_remover, mesmo com códigos repetidos
merged = pd.merge(df, df_remover[['Material']], on='Material', how='left', indicator=True)
# Filtra as linhas onde a coluna '_merge' é 'left_only' e remove a coluna '_merge'
df_resultado = merged[merged['_merge'] == 'left_only'].drop('_merge', axis=1)
df = df_resultado
# Criar uma Coluna com a data
data_atual = date.today().strftime('%d/%m/%Y')
df_resultado.insert(df_resultado.columns.get_loc('Tipo'), 'Mês Ano', data_atual)

#Somatório dos totais
df['Tipo'] = df['Tipo'].str.capitalize()
total = df.groupby('Tipo')['Val.total'].sum()
# Pula duas colunas a partir do último índice preenchido e insira os totais calculados
df.at[0, 'Custo Total'] = total.sum()
df.at[0,'Total MP'] = total.get('Materia prima', 0)
df.at[0, 'Total IMP'] = total.get('Improdutivo', 0)

# Define o caminho para o arquivo novo.xls
caminho_arquivo_novo = 'C:/Users/lcrodrigues/Documents/custo/novo.xls'

# Cria um novo arquivo Excel se não existir
if not os.path.exists(caminho_arquivo_novo):
    novo_workbook = Workbook()
    novo_workbook.save(caminho_arquivo_novo)

# Define o caminho para o arquivo novo.xlsx
caminho_xlsx = 'C:/Users/lcrodrigues/Documents/custo/novo.xlsx'

# Se o arquivo já existir, remove-o antes de salvar
if os.path.exists(caminho_xlsx):
    os.remove(caminho_xlsx)

# Salva o DataFrame df_resultado no arquivo novo.xlsx
df_resultado.to_excel(caminho_xlsx, index=False)