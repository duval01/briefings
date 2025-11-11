import streamlit as st
import pandas as pd
import requests
from io import StringIO
from urllib3.exceptions import InsecureRequestWarning
import os
from datetime import datetime
import io
import re

# --- CONFIGURAÇÕES GLOBAIS ---
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

MESES_MAPA = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
    "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
LISTA_MESES = list(MESES_MAPA.keys())

# Colunas necessárias
NCM_COLS = ['VL_FOB', 'CO_PAIS', 'CO_MES', 'SG_UF_NCM', 'CO_NCM']
NCM_DTYPES = {'CO_NCM': str, 'CO_SH4': str}

# --- FUNÇÕES DE LÓGICA (Helpers) ---

@st.cache_data(ttl=3600)
def ler_dados_csv_online(url, usecols=None, dtypes=None):
    """Lê dados CSV da URL com retentativas."""
    retries = 3
    for attempt in range(retries):
        try:
            resposta = requests.get(url, verify=False, timeout=(10, 1200)) 
            resposta.raise_for_status()
            df = pd.read_csv(StringIO(resposta.content.decode('latin-1')), encoding='latin-1',
                             sep=';', dtype=dtypes, usecols=usecols)
            return df
        except requests.exceptions.RequestException as e:
            st.error(f"Erro ao acessar o CSV (tentativa {attempt + 1}/{retries}): {e}")
            if "Read timed out" in str(e) and attempt < retries - 1:
                st.warning("Download demorou muito. Tentando novamente...")
                continue
            else:
                st.error(f"Falha ao baixar após {retries} tentativas.")
                return None
        except Exception as e:
            st.error(f"Erro inesperado ao baixar ou processar o CSV: {e}")
            return None
    return None

@st.cache_data(ttl=3600)
def carregar_dataframe(url, nome_arquivo, usecols=None, dtypes=None, mostrar_progresso=True):
    """Carrega o DataFrame da URL (usa cache) com colunas e dtypes."""
    progress_bar = None
    if mostrar_progresso: 
        progress_bar = st.progress(0, text=f"Carregando {nome_arquivo}...")
    
    df = ler_dados_csv_online(url, usecols=usecols, dtypes=dtypes)
    
    if mostrar_progresso and progress_bar: 
        if df is not None:
            progress_bar.progress(100, text=f"{nome_arquivo} carregado com sucesso.")
        else:
            progress_bar.empty()
    return df

@st.cache_data
def obter_dados_paises():
    """Carrega a tabela de países (ID e Nome) e armazena em cache."""
    url_pais = "https://balanca.economia.gov.br/balanca/bd/tabelas/PAIS.csv"
    df_pais = carregar_dataframe(url_pais, "PAIS.csv", usecols=['NO_PAIS', 'CO_PAIS'], mostrar_progresso=False) 
    if df_pais is not None and not df_pais.empty:
        # Cria um mapa de Código -> Nome
        return pd.Series(df_pais.NO_PAIS.values, index=df_pais.CO_PAIS).to_dict()
    return {}

@st.cache_data
def obter_lista_de_produtos_sh4():
    """Retorna uma lista de produtos (SH4)."""
    url_ncm = "https://balanca.economia.gov.br/balanca/bd/tabelas/NCM_SH.csv"
    df_ncm = carregar_dataframe(url_ncm, "NCM_SH.csv", usecols=['CO_SH4', 'NO_SH4_POR'], mostrar_progresso=False)
    if df_ncm is not None:
        # Formata como "Código - Nome"
        df_ncm = df_ncm.drop_duplicates(subset=['CO_SH4']).dropna() # Remove duplicados
        df_ncm['Display'] = df_ncm['CO_SH4'].astype(str) + " - " + df_ncm['NO_SH4_POR']
        lista_produtos = df_ncm['Display'].unique().tolist()
        lista_produtos.sort()
        return lista_produtos
    return ["Erro ao carregar lista de produtos"]

def get_sh4(co_ncm):
    """Extrai SH4 de um CO_NCM."""
    co_ncm_str = str(co_ncm)
    if pd.isna(co_ncm_str):
        return None
    if len(co_ncm_str) == 8:
        return co_ncm_str[:4]
    elif len(co_ncm_str) >= 7:
        return co_ncm_str[:3]
    else:
        return None

def formatar_valor(valor):
    prefixo = ""
    if valor < 0:
        prefixo = "-"
        valor = abs(valor)

    if valor >= 1_000_000_000:
        valor_formatado_str = f"{(valor / 1_000_000_000):.2f}".replace('.',',')
        unidade = "bilhão" if (valor / 1_000_000_000) < 2 else "bilhões"
        return f"{prefixo}US$ {valor_formatado_str} {unidade}"
    if valor >= 1_000_000:
        valor_formatado_str = f"{(valor / 1_000_000):.2f}".replace('.',',')
        unidade = "milhão" if (valor / 1_000_000) < 2 else "milhões"
        return f"{prefixo}US$ {valor_formatado_str} {unidade}"
    if valor >= 1_000:
        valor_formatado_str = f"{(valor / 1_000):.2f}".replace('.',',')
        return f"{prefixo}US$ {valor_formatado_str} mil"
    
    valor_formatado_str = f"{valor:.2f}".replace('.',',')
    return f"{prefixo}US$ {valor_formatado_str}"

# --- CONFIGURAÇÃO DA PÁGINA ---

# st.set_page_config(page_title="Análise por Produto", layout="wide") # Config é feito no app.py
st.sidebar.empty()
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

st.header("1. Configurações da Análise de Produto (NCM)")

lista_de_produtos = obter_lista_de_produtos_sh4()
mapa_nomes_paises = obter_dados_paises()
ano_atual = datetime.now().year

col1, col2 = st.columns(2)
with col1:
    ano_principal = st.number_input(
        "Ano de Referência:", min_value=1998, max_value=ano_atual, value=ano_atual,
        help="O ano principal que você quer analisar."
    )
    produtos_selecionados = st.multiselect(
        "Selecione o(s) produto(s) (SH4):",
        options=lista_de_produtos,
        default=[p for p in lista_de_produtos if "0901" in p][:1], # Pega "0901 - Café"
        help="Você pode digitar para pesquisar. O filtro usa os 4 dígitos do SH4."
    )

with col2:
    ano_comparacao = st.number_input(
        "Ano de Comparação:", min_value=1998, max_value=ano_atual, value=ano_atual - 1,
        help="O ano contra o qual você quer comparar."
    )
    meses_selecionados = st.multiselect(
        "Meses de Análise (opcional):",
        options=LISTA_MESES,
        help="Selecione os meses. Se deixar em branco, o ano inteiro será analisado."
    )

st.header("2. Gerar Análise")

if st.button("Iniciar Análise por Produto"):
    
    with st.spinner(f"Processando dados de produto..."):
        try:
            # --- Validação ---
            if not produtos_selecionados:
                st.error("Nenhum produto selecionado.")
                st.stop()
            # Extrai apenas os códigos SH4
            codigos_sh4_selecionados = [s.split(" - ")[0] for s in produtos_selecionados]
            
            # --- URLs ---
            url_exp_ano_principal = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/ncm/EXP_{ano_principal}.csv"
            url_exp_ano_comparacao = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/ncm/EXP_{ano_comparacao}.csv"
            url_imp_ano_principal = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/ncm/IMP_{ano_principal}.csv"
            url_imp_ano_comparacao = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/ncm/IMP_{ano_comparacao}.csv"

            # --- Carregamento ---
            df_exp_princ = carregar_dataframe(url_exp_ano_principal, f"EXP_{ano_principal}.csv", usecols=NCM_COLS, dtypes=NCM_DTYPES)
            df_exp_comp = carregar_dataframe(url_exp_ano_comparacao, f"EXP_{ano_comparacao}.csv", usecols=NCM_COLS, dtypes=NCM_DTYPES)
            df_imp_princ = carregar_dataframe(url_imp_ano_principal, f"IMP_{ano_principal}.csv", usecols=NCM_COLS, dtypes=NCM_DTYPES)
            df_imp_comp = carregar_dataframe(url_imp_ano_comparacao, f"IMP_{ano_comparacao}.csv", usecols=NCM_COLS, dtypes=NCM_DTYPES)

            if df_exp_princ is None or df_imp_princ is None or df_exp_comp is None or df_imp_comp is None:
                st.error("Falha ao carregar arquivos de dados NCM. Tente novamente.")
                st.stop()
            
            # --- Filtro de Meses ---
            if meses_selecionados:
                meses_para_filtrar = [MESES_MAPA[m] for m in meses_selecionados]
            else:
                meses_para_filtrar = list(range(1, df_exp_princ['CO_MES'].max() + 1))
            
            # --- Adiciona coluna SH4 e Filtra ---
            st.subheader(f"Análise de: {', '.join(produtos_selecionados)}")
            
            df_exp_princ['SH4'] = df_exp_princ['CO_NCM'].apply(get_sh4)
            df_exp_comp['SH4'] = df_exp_comp['CO_NCM'].apply(get_sh4)
            df_imp_princ['SH4'] = df_imp_princ['CO_NCM'].apply(get_sh4)
            df_imp_comp['SH4'] = df_imp_comp['CO_NCM'].apply(get_sh4)
            
            # --- Processamento Exportação ---
            st.header("Principais Destinos (Exportação de MG)")
            
            df_exp_princ_f = df_exp_princ[(df_exp_princ['SG_UF_NCM'] == 'MG') & (df_exp_princ['SH4'].isin(codigos_sh4_selecionados)) & (df_exp_princ['CO_MES'].isin(meses_para_filtrar))]
            df_exp_comp_f = df_exp_comp[(df_exp_comp['SG_UF_NCM'] == 'MG') & (df_exp_comp['SH4'].isin(codigos_sh4_selecionados)) & (df_exp_comp['CO_MES'].isin(meses_para_filtrar))]
            
            exp_paises_princ = df_exp_princ_f.groupby('CO_PAIS')['VL_FOB'].sum().sort_values(ascending=False).reset_index()
            exp_paises_comp = df_exp_comp_f.groupby('CO_PAIS')['VL_FOB'].sum().reset_index()

            exp_paises_princ['País'] = exp_paises_princ['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            exp_paises_princ[f'Valor {ano_principal} (US$)'] = exp_paises_princ['VL_FOB']
            
            exp_paises_comp['País'] = exp_paises_comp['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            exp_paises_comp[f'Valor {ano_comparacao} (US$)'] = exp_paises_comp['VL_FOB']
            
            exp_final = pd.merge(exp_paises_princ[['País', f'Valor {ano_principal} (US$)']], 
                                 exp_paises_comp[['País', f'Valor {ano_comparacao} (US$)']], 
                                 on="País", how="outer").fillna(0)
            
            exp_final['Variação %'] = 100 * (exp_final[f'Valor {ano_principal} (US$)'] - exp_final[f'Valor {ano_comparacao} (US$)']) / exp_final[f'Valor {ano_comparacao} (US$)']
            exp_final['Variação %'] = exp_final['Variação %'].fillna(0).round(2)

            exp_final[f'Valor {ano_principal}'] = exp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
            exp_final[f'Valor {ano_comparacao}'] = exp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)
            
            st.dataframe(exp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).head(20)
                         [['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']])
            
            del df_exp_princ, df_exp_comp, df_exp_princ_f, df_exp_comp_f, exp_paises_princ, exp_paises_comp, exp_final

            # --- Processamento Importação ---
            st.header("Principais Origens (Importação de MG)")
            
            df_imp_princ_f = df_imp_princ[(df_imp_princ['SG_UF_NCM'] == 'MG') & (df_imp_princ['SH4'].isin(codigos_sh4_selecionados)) & (df_imp_princ['CO_MES'].isin(meses_para_filtrar))]
            df_imp_comp_f = df_imp_comp[(df_imp_comp['SG_UF_NCM'] == 'MG') & (df_imp_comp['SH4'].isin(codigos_sh4_selecionados)) & (df_imp_comp['CO_MES'].isin(meses_para_filtrar))]

            imp_paises_princ = df_imp_princ_f.groupby('CO_PAIS')['VL_FOB'].sum().sort_values(ascending=False).reset_index()
            imp_paises_comp = df_imp_comp_f.groupby('CO_PAIS')['VL_FOB'].sum().reset_index()

            imp_paises_princ['País'] = imp_paises_princ['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            imp_paises_princ[f'Valor {ano_principal} (US$)'] = imp_paises_princ['VL_FOB']
            
            imp_paises_comp['País'] = imp_paises_comp['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            imp_paises_comp[f'Valor {ano_comparacao} (US$)'] = imp_paises_comp['VL_FOB']
            
            imp_final = pd.merge(imp_paises_princ[['País', f'Valor {ano_principal} (US$)']], 
                                 imp_paises_comp[['País', f'Valor {ano_comparacao} (US$)']], 
                                 on="País", how="outer").fillna(0)

            imp_final['Variação %'] = 100 * (imp_final[f'Valor {ano_principal} (US$)'] - imp_final[f'Valor {ano_comparacao} (US$)']) / imp_final[f'Valor {ano_comparacao} (US$)']
            imp_final['Variação %'] = imp_final['Variação %'].fillna(0).round(2)
            
            imp_final[f'Valor {ano_principal}'] = imp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
            imp_final[f'Valor {ano_comparacao}'] = imp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)

            st.dataframe(imp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).head(20)
                         [['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']])
            
            del df_imp_princ, df_imp_comp, df_imp_princ_f, df_imp_comp_f, imp_paises_princ, imp_paises_comp, imp_final

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante a análise de produto:")
            st.exception(e)
