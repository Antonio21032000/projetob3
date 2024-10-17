import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

def clean_volume(value):
    if pd.isna(value):
        return np.nan
    # Remove 'R$', vírgulas e espaços, depois converte para float
    cleaned = str(value).replace('R$', '').replace(',', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return np.nan

# Leitura do CSV
tabela_diretoria = pd.read_csv(r'M:\VS Code\teste.csv', encoding='latin1', sep=';')

tabela_diretoria.to_excel(r'M:\VS Code\teste1.xlsx')

# Imprime as colunas para debug
print("Colunas no DataFrame:")
print(tabela_diretoria.columns)

# Tenta identificar a coluna de volume financeiro
volume_cols = [col for col in tabela_diretoria.columns if 'volume' in col.lower()]

if volume_cols:
    volume_col = volume_cols[0]
    print(f"Coluna de volume financeiro identificada: {volume_col}")
    
    # Limpeza da coluna de Volume Financeiro
    tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(clean_volume)
    
    # Removendo as colunas especificadas
    colunas_para_remover = ['Nome_Companhia', 'Intermediario', 'Versao']
    tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])
    
    # Removendo duplicatas baseadas na coluna de Volume Financeiro, mantendo a primeira ocorrência
    tabela_diretoria = tabela_diretoria.drop_duplicates(subset=[volume_col], keep='first')
    
    # Ordenando o DataFrame pelo Volume Financeiro em ordem decrescente
    tabela_diretoria = tabela_diretoria.sort_values(by=volume_col, ascending=False)
    
    # Formatando a coluna de Volume Financeiro
    tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
    
    # Formatando a coluna "Quantidade"
    if 'Quantidade' in tabela_diretoria.columns:
        tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
    
    # Formatando a coluna "Preco_Unitario"
    if 'Preco_Unitario' in tabela_diretoria.columns:
        tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'{x:.2f}' if pd.notnull(x) else '')
    
    print(tabela_diretoria)
    print(tabela_diretoria.info())
    
    # Verificação adicional
    print(f"\nValores únicos na coluna {volume_col}:")
    print(tabela_diretoria[volume_col].unique())
else:
    print("Não foi possível identificar uma coluna de volume financeiro.")
    print("Por favor, forneça o nome exato da coluna que contém os valores de volume financeiro.")

# Opcional: Salvar o DataFrame limpo em um novo arquivo CSV
# tabela_diretoria.to_csv('tabela_diretoria_limpa.csv', index=False)

# Após processar a tabela_diretoria e antes de gerar o Excel, adicione:
tabela_diretoria.to_pickle('tabela_diretoria.pkl')

# Adicione o código Streamlit aqui
st.title('Tabela Diretoria')

# Exibir a tabela
st.dataframe(tabela_diretoria)

# Adicionar alguns filtros básicos
st.sidebar.header('Filtros')

# Filtro de data
if 'Data_Operacao' in tabela_diretoria.columns:
    min_date = tabela_diretoria['Data_Operacao'].min()
    max_date = tabela_diretoria['Data_Operacao'].max()
    date_range = st.sidebar.date_input('Intervalo de Datas', [min_date, max_date])

# Filtro de empresa
if 'Empresa' in tabela_diretoria.columns:
    empresas = st.sidebar.multiselect('Empresas', options=tabela_diretoria['Empresa'].unique())

# Filtro de tipo de operação
if 'Tipo_Operacao' in tabela_diretoria.columns:
    tipos_operacao = st.sidebar.multiselect('Tipos de Operação', options=tabela_diretoria['Tipo_Operacao'].unique())

# Aplicar filtros
filtered_df = tabela_diretoria.copy()

if 'Data_Operacao' in tabela_diretoria.columns and len(date_range) == 2:
    filtered_df = filtered_df[(filtered_df['Data_Operacao'] >= date_range[0]) & 
                              (filtered_df['Data_Operacao'] <= date_range[1])]

if empresas:
    filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]

if tipos_operacao:
    filtered_df = filtered_df[filtered_df['Tipo_Operacao'].isin(tipos_operacao)]

# Exibir a tabela filtrada
st.dataframe(filtered_df)

# Gerar arquivo Excel
excel_path = r'M:\VS Code\tabela_diretoria.xlsx'

# Criando um ExcelWriter object
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    tabela_diretoria.to_excel(writer, index=False, sheet_name='Dados')
    
    # Obtendo a planilha ativa
    workbook = writer.book
    worksheet = workbook['Dados']
    
    # Ajustando a largura das colunas
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print(f"Arquivo Excel gerado com sucesso: {excel_path}")


