import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st

# Configuração da página Streamlit
st.set_page_config(layout="wide", page_title="Dashboard STK")

# Cores da STK
STK_COLORS = {
    'primary': '#102E46',  # Azul escuro
    'secondary': '#C98C2E',  # Dourado
    'accent': '#0E7C7B',  # Cor adicional (turquesa)
    'background': '#FFFFFF',  # Branco para o fundo
    'text': '#333333',  # Cinza escuro para texto
}

# Aplicar estilos CSS personalizados
st.markdown(f"""
    <style>
    .reportview-container .main .block-container{{
        max-width: 1200px;
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 5rem;
        padding-right: 5rem;
    }}
    .stApp {{
        background: linear-gradient(135deg, {STK_COLORS['primary']}, {STK_COLORS['secondary']});
        color: {STK_COLORS['text']};
    }}
    .stButton>button {{
        color: white;
        background-color: {STK_COLORS['accent']};
        border-radius: 5px;
    }}
    .stSelectbox, .stMultiSelect {{
        background-color: rgba(255, 255, 255, 0.8);
        color: {STK_COLORS['text']};
    }}
    h1 {{
        color: white;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
    }}
    .stDateInput>div>div>input {{
        color: {STK_COLORS['text']};
        background-color: rgba(255, 255, 255, 0.8);
    }}
    .stDataFrame {{
        background-color: rgba(255, 255, 255, 0.9);
    }}
    </style>
    """, unsafe_allow_html=True)

# Função para limpar o volume financeiro
def clean_volume(value):
    if pd.isna(value):
        return np.nan
    cleaned = str(value).replace('R$', '').replace(',', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return np.nan

# Leitura do CSV
@st.cache_data
def load_data():
    df = pd.read_csv('teste.csv', encoding='latin1', sep=';')
    return df

tabela_diretoria = load_data()

# Processamento dos dados
volume_cols = [col for col in tabela_diretoria.columns if 'volume' in col.lower()]

if volume_cols:
    volume_col = volume_cols[0]
    tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(clean_volume)
    
    colunas_para_remover = ['Nome_Companhia', 'Intermediario', 'Versao']
    tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])
    
    tabela_diretoria = tabela_diretoria.drop_duplicates(subset=[volume_col], keep='first')
    tabela_diretoria = tabela_diretoria.sort_values(by=volume_col, ascending=False)
    
    tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
    
    if 'Quantidade' in tabela_diretoria.columns:
        tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
    
    if 'Preco_Unitario' in tabela_diretoria.columns:
        tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'{x:.2f}' if pd.notnull(x) else '')

# Interface Streamlit
st.title('Dashboard STK')

# Filtros
col1, col2 = st.columns(2)

with col1:
    empresas = st.multiselect('Empresas', options=sorted(tabela_diretoria['Empresa'].unique()))

with col2:
    if 'Data_Referencia' in tabela_diretoria.columns:
        tabela_diretoria['Data_Referencia'] = pd.to_datetime(tabela_diretoria['Data_Referencia'])
        min_date = tabela_diretoria['Data_Referencia'].min().date()
        max_date = tabela_diretoria['Data_Referencia'].max().date()
        date_range = st.date_input('Intervalo de Datas', [min_date, max_date])

# Aplicar filtros
filtered_df = tabela_diretoria.copy()

if empresas:
    filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]

if 'Data_Referencia' in tabela_diretoria.columns and len(date_range) == 2:
    filtered_df = filtered_df[(filtered_df['Data_Referencia'].dt.date >= date_range[0]) & 
                              (filtered_df['Data_Referencia'].dt.date <= date_range[1])]

# Exibir a tabela filtrada
st.dataframe(filtered_df, use_container_width=True, height=600)

# Gerar arquivo Excel (mantido do código original, mas sem a mensagem de sucesso)
excel_path = 'tabela_diretoria.xlsx'

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    tabela_diretoria.to_excel(writer, index=False, sheet_name='Dados')
    
    workbook = writer.book
    worksheet = workbook['Dados']
    
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



