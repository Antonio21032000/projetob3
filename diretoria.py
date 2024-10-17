import pandas as pd
import numpy as np
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
        font-weight: bold !important;
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
tabela_diretoria['Empresa'] = tabela_diretoria['Empresa'].astype(str)

# Renomear a coluna de volume
volume_col = 'Volume,,,,'  # Nome original da coluna
if volume_col in tabela_diretoria.columns:
    tabela_diretoria[volume_col] = tabela_diretoria[volume_col].apply(clean_volume)
    tabela_diretoria.rename(columns={volume_col: 'Volume Financeiro (R$)'}, inplace=True)
    
    tabela_diretoria = tabela_diretoria.sort_values(by='Volume Financeiro (R$)', ascending=False)
    
    tabela_diretoria['Volume Financeiro (R$)'] = tabela_diretoria['Volume Financeiro (R$)'].apply(lambda x: f'R$ {x:,.2f}' if pd.notnull(x) else '')

# Remover colunas específicas
colunas_para_remover = ['CNPJ_Companhia', 'Tipo_Empresa', 'Descricao_Movimentacao', 'Tipo_Operacao', 'Nome_Companhia', 'Intermediario', 'Versao']
tabela_diretoria = tabela_diretoria.drop(columns=[col for col in colunas_para_remover if col in tabela_diretoria.columns])

if 'Quantidade' in tabela_diretoria.columns:
    tabela_diretoria['Quantidade'] = tabela_diretoria['Quantidade'].apply(lambda x: f'{x:,.0f}' if pd.notnull(x) else '')

if 'Preco_Unitario' in tabela_diretoria.columns:
    tabela_diretoria['Preco_Unitario'] = tabela_diretoria['Preco_Unitario'].apply(lambda x: f'R$ {x:.2f}' if pd.notnull(x) else '')

# Interface Streamlit
st.title('Dashboard STK')

# Filtros
col1, col2, col3, col4 = st.columns(4)

with col1:
    empresas_unicas = sorted(tabela_diretoria['Empresa'].dropna().unique())
    empresas = st.multiselect('Empresas', options=empresas_unicas)

with col2:
    if 'Data_Referencia' in tabela_diretoria.columns:
        tabela_diretoria['Data_Referencia'] = pd.to_datetime(tabela_diretoria['Data_Referencia'])
        min_date = tabela_diretoria['Data_Referencia'].min().date()
        max_date = tabela_diretoria['Data_Referencia'].max().date()
        date_range = st.date_input('Intervalo de Datas', [min_date, max_date])

with col3:
    tipos_movimentacao = sorted(tabela_diretoria['Tipo_Movimentacao'].dropna().unique())
    tipos_movimentacao_selecionados = st.multiselect('Tipo de Movimentação', options=tipos_movimentacao)

with col4:
    tipos_cargo = sorted(tabela_diretoria['Tipo_Cargo'].dropna().unique())
    tipos_cargo_selecionados = st.multiselect('Tipo de Cargo', options=tipos_cargo)

# Aplicar filtros
filtered_df = tabela_diretoria.copy()

if empresas:
    filtered_df = filtered_df[filtered_df['Empresa'].isin(empresas)]

if 'Data_Referencia' in tabela_diretoria.columns and len(date_range) == 2:
    filtered_df = filtered_df[(filtered_df['Data_Referencia'].dt.date >= date_range[0]) & 
                              (filtered_df['Data_Referencia'].dt.date <= date_range[1])]

if tipos_movimentacao_selecionados:
    filtered_df = filtered_df[filtered_df['Tipo_Movimentacao'].isin(tipos_movimentacao_selecionados)]

if tipos_cargo_selecionados:
    filtered_df = filtered_df[filtered_df['Tipo_Cargo'].isin(tipos_cargo_selecionados)]

# Exibir a tabela filtrada
st.dataframe(filtered_df, use_container_width=True, height=600)

# Gerar arquivo Excel
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
