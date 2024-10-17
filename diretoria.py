import pandas as pd
import numpy as np
import streamlit as st

# ... [Código de configuração e estilo anterior permanece o mesmo] ...

# Adicione este estilo CSS para a primeira linha da tabela
st.markdown("""
    <style>
    .stDataFrame thead tr th {
        font-weight: bold !important;
    }
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
    
    tabela_diretoria = tabela_diretoria.sort_values(by='Volume Financeiro (R$)', ascending=False)
    
    tabela_diretoria['Volume Financeiro (R$)'] = tabela_diretoria['Volume Financeiro (R$)'].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')
    
    if 'Quantidade' in tabela_diretoria.columns:
        tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')
    
    if 'Preco_Unitario' in tabela_diretoria.columns:
        tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'R$ {x:.2f}' if pd.notnull(x) else '')

# Interface Streamlit
st.title('Dashboard STK')

# Filtros
col1, col2, col3, col4 = st.columns(4)

with col1:
    # Corrigindo o filtro de empresas para incluir todas as empresas únicas
    empresas = st.multiselect('Empresas', options=sorted(tabela_diretoria['Empresa'].unique()))

with col2:
    if 'Data_Referencia' in tabela_diretoria.columns:
        tabela_diretoria['Data_Referencia'] = pd.to_datetime(tabela_diretoria['Data_Referencia'])
        min_date = tabela_diretoria['Data_Referencia'].min().date()
        max_date = tabela_diretoria['Data_Referencia'].max().date()
        date_range = st.date_input('Intervalo de Datas', [min_date, max_date])

with col3:
    tipos_movimentacao = st.multiselect('Tipo de Movimentação', options=sorted(tabela_diretoria['Tipo_Movimentacao'].unique()))

with col4:
    tipos_cargo = st.multiselect('Tipo de Cargo', options=sorted(tabela_diretoria['Tipo_Cargo'].unique()))

# Aplicar filtros
filtered_df = tabela_diretoria.copy()

if empresas:
    filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]

if 'Data_Referencia' in tabela_diretoria.columns and len(date_range) == 2:
    filtered_df = filtered_df[(filtered_df['Data_Referencia'].dt.date >= date_range[0]) & 
                              (filtered_df['Data_Referencia'].dt.date <= date_range[1])]

if tipos_movimentacao:
    filtered_df = filtered_df[filtered_df['Tipo_Movimentacao'].isin(tipos_movimentacao)]

if tipos_cargo:
    filtered_df = filtered_df[filtered_df['Tipo_Cargo'].isin(tipos_cargo)]

# Exibir a tabela filtrada
st.dataframe(filtered_df, use_container_width=True, height=600)

# ... [O restante do código permanece o mesmo] ...
