import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st
import base64
from io import BytesIO

# Configuração da página Streamlit
st.set_page_config(layout="wide", page_title="Tracker of Insiders")

# Cores da STK
STK_COLORS = {
    'primary': '#102F46',  # Azul escuro
    'secondary': '#C9B22E',  # Dourado
    'accent': '#0990B2',  # Azul claro
    'background': '#F5F7FA',  # Cinza muito claro para o fundo
    'text': '#081824',  # Azul muito escuro para texto
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
        background: linear-gradient(135deg, {STK_COLORS['primary']}, {STK_COLORS['accent']});
        color: {STK_COLORS['text']};
    }}
    .stButton>button {{
        color: white;
        background-color: {STK_COLORS['accent']};
        border-radius: 5px;
        font-weight: bold;
        border: none;
        padding: 0.5rem 1rem;
        transition: background-color 0.3s;
    }}
    .stButton>button:hover {{
        background-color: {STK_COLORS['primary']};
    }}
    .stSelectbox, .stMultiSelect {{
        background-color: white;
        border: 1px solid {STK_COLORS['accent']};
        border-radius: 5px;
        color: {STK_COLORS['text']};
    }}
    h1 {{
        color: white;
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 2rem;
        padding: 1rem;
        background: linear-gradient(90deg, {STK_COLORS['primary']}, {STK_COLORS['accent']});
        border-radius: 10px;
    }}
    .stDateInput>div>div>input {{
        color: {STK_COLORS['text']};
        background-color: white;
        border: 1px solid {STK_COLORS['accent']};
        border-radius: 5px;
    }}
    .stDataFrame {{
        background-color: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }}
    .stDataFrame table {{
        color: {STK_COLORS['text']} !important;
    }}
    .stDataFrame th {{
        background-color: {STK_COLORS['primary']} !important;
        color: white !important;
        padding: 0.5rem !important;
    }}
    .stDataFrame td {{
        background-color: white !important;
        padding: 0.5rem !important;
    }}
    .stDataFrame tr:nth-of-type(even) {{
        background-color: {STK_COLORS['background']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# Título
st.title('Tracker of Insiders')

# Função para limpar o volume financeiro
def clean_volume(value):
    if pd.isna(value):
        return np.nan
    cleaned = str(value).replace('R$', '').replace(',', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return np.nan

# Função para gerar link de download do Excel
def get_table_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="tabela_diretoria.xlsx">Download Excel file</a>'
    return href

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
    
    # Renomear a coluna de volume
    tabela_diretoria.rename(columns={volume_col: 'Volume Financeiro (R$)'}, inplace=True)
    
    # Remover colunas específicas
    colunas_para_remover = ['CNPJ_Companhia', 'Tipo_Empresa', 'Descricao_Movimentacao', 'Tipo_Operacao', 'Nome_Companhia', 'Intermediario', 'Versao']
    tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])
    
    tabela_diretoria = tabela_diretoria.drop_duplicates(subset=['Volume Financeiro (R$)'], keep='first')
    tabela_diretoria = tabela_diretoria.sort_values(by='Volume Financeiro (R$)', ascending=False)
    
    tabela_diretoria['Volume Financeiro (R$)'] = tabela_diretoria['Volume Financeiro (R$)'].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
    
    if 'Quantidade' in tabela_diretoria.columns:
        tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
    
    if 'Preco_Unitario' in tabela_diretoria.columns:
        tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'R$ {x:.2f}' if pd.notnull(x) else '')

# Filtros
col1, col2 = st.columns(2)

with col1:
    empresas = st.multiselect('Empresas', options=sorted(tabela_diretoria['Empresa'].unique()), key="empresas_select")

with col2:
    if 'Data_Referencia' in tabela_diretoria.columns:
        tabela_diretoria['Data_Referencia'] = pd.to_datetime(tabela_diretoria['Data_Referencia'])
        min_date = tabela_diretoria['Data_Referencia'].min().date()
        max_date = tabela_diretoria['Data_Referencia'].max().date()
        date_range = st.date_input('Intervalo de Datas', [min_date, max_date], key="date_range")

# Aplicar filtros
filtered_df = tabela_diretoria.copy()

if empresas:
    filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]

if 'Data_Referencia' in tabela_diretoria.columns and len(date_range) == 2:
    filtered_df = filtered_df[(filtered_df['Data_Referencia'].dt.date >= date_range[0]) & 
                              (filtered_df['Data_Referencia'].dt.date <= date_range[1])]

# Exibir a tabela filtrada
st.dataframe(filtered_df.reset_index(drop=True), use_container_width=True, height=600)

# Botão para download do Excel
st.markdown(get_table_download_link(filtered_df), unsafe_allow_html=True)
