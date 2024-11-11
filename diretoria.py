import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st
import base64
from io import BytesIO

# Configuração da página Streamlit
st.set_page_config(layout="wide", page_title="Análise de Insiders")

# Cores atualizadas
BG_COLOR = '#102F46'
TITLE_BG_COLOR = '#DAA657'
TITLE_TEXT_COLOR = 'white'
TEXT_COLOR = '#333333'

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
        background-color: {BG_COLOR};
    }}
    .stButton>button {{
        color: {TITLE_BG_COLOR};
        background-color: white;
        border-radius: 5px;
        font-weight: bold;
        border: none;
        padding: 0.5rem 1rem;
        transition: background-color 0.3s;
    }}
    .stButton>button:hover {{
        background-color: #f0f0f0;
    }}
    .stSelectbox, .stMultiSelect {{
        background-color: white;
        border-radius: 5px;
        color: {TEXT_COLOR};
    }}
    .title-container {{
        background-color: {TITLE_BG_COLOR};
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }}
    .title-container h1 {{
        color: {TITLE_TEXT_COLOR};
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin: 0;
    }}
    .stDateInput>div>div>input {{
        color: {TEXT_COLOR};
        background-color: white;
        border-radius: 5px;
    }}
    .stDataFrame {{
        background-color: white;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }}
    .stDataFrame table {{
        color: {TEXT_COLOR} !important;
    }}
    .stDataFrame th {{
        background-color: {TITLE_BG_COLOR} !important;
        color: {TITLE_TEXT_COLOR} !important;
        padding: 0.5rem !important;
    }}
    .stDataFrame td {{
        background-color: white !important;
        padding: 0.5rem !important;
    }}
    .stDataFrame tr:nth-of-type(even) {{
        background-color: #f8f8f8 !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# Título
st.markdown('<div class="title-container"><h1>Análise de Insiders BR</h1></div>', unsafe_allow_html=True)

# Função para limpar volume financeiro
def clean_volume(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace('R$', '').replace('.', '').replace(',', '.').replace(' ', '').strip()
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
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="tabela_diretoria.xlsx">Baixar arquivo Excel</a>'
    return href

# Ler CSV
@st.cache_data(ttl=3600)
def load_data():
    df = pd.read_csv('teste.csv', encoding='latin1', sep=';')
    # Converter a coluna de data para datetime logo após a leitura
    if 'Data_Referencia' in df.columns:
        df['Data_Referencia'] = pd.to_datetime(df['Data_Referencia'], format='%Y-%m-%d', errors='coerce')
    return df

try:
    tabela_diretoria = load_data()
    
    # Processamento dos dados
    volume_cols = [col for col in tabela_diretoria.columns if 'volume' in col.lower()]
    
    if volume_cols:
        volume_col = volume_cols[0]
        # Converter volume para float
        tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(clean_volume)
        
        # Renomear coluna de volume
        tabela_diretoria.rename(columns={volume_col: 'Volume Financeiro (R$)'}, inplace=True)
        
        # Remover colunas específicas
        colunas_para_remover = ['CNPJ_Companhia', 'Tipo_Empresa', 'Descricao_Movimentacao', 'Tipo_Operacao', 'Nome_Companhia', 'Intermediario', 'Versao']
        tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])
        
        # Remover duplicatas e ordenar
        tabela_diretoria = tabela_diretoria.drop_duplicates(subset=['Volume Financeiro (R$)'], keep='first')
        tabela_diretoria = tabela_diretoria.sort_values(by='Data_Referencia', ascending=False)
        
        # Converter valores numéricos antes de formatar
        if 'Quantidade' in tabela_diretoria.columns:
            tabela_diretoria['Quantidade'] = pd.to_numeric(tabela_diretoria['Quantidade'], errors='coerce')
        
        if 'Preco_Unitario' in tabela_diretoria.columns:
            tabela_diretoria['Preco_Unitario'] = pd.to_numeric(tabela_diretoria['Preco_Unitario'], errors='coerce')
        
        # Formatar valores para exibição
        tabela_diretoria['Volume Financeiro (R$)'] = tabela_diretoria['Volume Financeiro (R$)'].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
        
        if 'Quantidade' in tabela_diretoria.columns:
            tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
        
        if 'Preco_Unitario' in tabela_diretoria.columns:
            tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'R$ {x:.2f}' if pd.notnull(x) else '')
    
    # Filtros
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if 'Empresa' in tabela_diretoria.columns:
            empresas = st.multiselect('Empresas', options=sorted(tabela_diretoria['Empresa'].unique()), key="empresas_select")
    
    with col2:
        if 'Data_Referencia' in tabela_diretoria.columns:
            min_date = tabela_diretoria['Data_Referencia'].min()
            max_date = tabela_diretoria['Data_Referencia'].max()
            date_range = st.date_input(
                'Intervalo de Datas',
                [min_date.date(), max_date.date()],
                min_value=min_date.date(),
                max_value=max_date.date(),
                key="date_range"
            )
    
    with col3:
        if 'Tipo_Movimentacao' in tabela_diretoria.columns:
            tipo_movimentacao = st.multiselect('Tipo de Movimentação', options=sorted(tabela_diretoria['Tipo_Movimentacao'].unique()), key="tipo_movimentacao_select")
    
    # Aplicar filtros
    filtered_df = tabela_diretoria.copy()
    
    if empresas:
        filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]
    
    if len(date_range) == 2:
        filtered_df = filtered_df[
            (filtered_df['Data_Referencia'].dt.date >= date_range[0]) & 
            (filtered_df['Data_Referencia'].dt.date <= date_range[1])
        ]
    
    if tipo_movimentacao:
        filtered_df = filtered_df[filtered_df['Tipo_Movimentacao'].isin(tipo_movimentacao)]
    
    # Exibir tabela filtrada
    st.dataframe(filtered_df.reset_index(drop=True), use_container_width=True, height=600)
    
    # Botão de download Excel
    st.markdown(get_table_download_link(filtered_df), unsafe_allow_html=True)

except Exception as e:
    st.error(f"Erro ao carregar ou processar os dados: {str(e)}")
    st.write("Por favor, verifique se o arquivo 'teste.csv' existe e está no formato correto.")
