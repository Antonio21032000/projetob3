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
    'text': '#FFFFFF',  # Branco para o texto
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
    }}
    .stButton>button {{
        color: white;
        background-color: {STK_COLORS['accent']};
        border-radius: 5px;
    }}
    .stSelectbox, .stMultiSelect {{
        background-color: rgba(255, 255, 255, 0.1);
        color: {STK_COLORS['text']};
    }}
    h1, h2, h3, p, label {{
        color: {STK_COLORS['text']};
    }}
    .stDateInput>div>div>input {{
        color: {STK_COLORS['text']};
        background-color: rgba(255, 255, 255, 0.1);
    }}
    .stDataFrame {{
        color: {STK_COLORS['text']};
    }}
    .stDataFrame table {{
        color: {STK_COLORS['text']} !important;
    }}
    .stDataFrame th {{
        background-color: {STK_COLORS['primary']} !important;
        color: {STK_COLORS['text']} !important;
    }}
    .stDataFrame td {{
        background-color: rgba(255, 255, 255, 0.1) !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# Função para limpar o volume financeiro
def clean_volume(value):
    if pd.isna(value):
        return np.nan
    cleaned = str(value).replace('R$', '').replace('.', '').replace(',', '.').strip()
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
    
    # Renomear a coluna de volume
    tabela_diretoria.rename(columns={volume_col: 'Volume Financeiro (R$)'}, inplace=True)
    
    # Remover colunas específicas
    colunas_para_remover = ['CNPJ_Companhia', 'Tipo_Empresa', 'Descricao_Movimentacao', 'Tipo_Operacao', 'Nome_Companhia', 'Intermediario', 'Versao']
    tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])
    
    tabela_diretoria = tabela_diretoria.drop_duplicates(subset=['Volume Financeiro (R$)'], keep='first')
    tabela_diretoria = tabela_diretoria.sort_values(by='Volume Financeiro (R$)', ascending=False)
    
    tabela_diretoria['Volume Financeiro (R$)'] = tabela_diretoria['Volume Financeiro (R$)'].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
    
    if 'Quantidade' in tabela_diretoria.columns:
        tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x:


