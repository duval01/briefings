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
MUN_COLS = ['VL_FOB', 'CO_PAIS', 'CO_MES', 'SG_UF_MUN', 'CO_MUN']
MUN_DTYPES = {'CO_MUN': str}

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
def obter_lista_de_municipios():
    """Retorna uma lista de nomes de municípios de MG."""
    url_uf_mun = "https://balanca.economia.gov.br/balanca/bd/tabelas/UF_MUN.csv"
    df_mun = carregar_dataframe(url_uf_mun, "UF_MUN.csv", usecols=['SG_UF', 'NO_MUN', 'CO_MUN_GEO'], mostrar_progresso=False)
    if df_mun is not None:
        lista_mun = df_mun[df_mun['SG_UF'] == 'MG']['NO_MUN'].unique().tolist()
        lista_mun.sort()
        return lista_mun
    return ["Erro ao carregar lista de municípios"]

@st.cache_data
def obter_mapa_codigos_municipios():
    """Retorna um mapa de Nome -> Código (CO_MUN_GEO) para municípios de MG."""
    url_uf_mun = "https://balanca.economia.gov.br/balanca/bd/tabelas/UF_MUN.csv"
    df_mun = carregar_dataframe(url_uf_mun, "UF_MUN.csv", usecols=['SG_UF', 'NO_MUN', 'CO_MUN_GEO'], mostrar_progresso=False)
    if df_mun is not None:
        df_mun_mg = df_mun[df_mun['SG_UF'] == 'MG']
        # Mapeia Nome para CO_MUN_GEO
        return pd.Series(df_mun_mg.CO_MUN_GEO.values, index=df_mun_mg.NO_MUN).to_dict()
    return {}

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

# st.set_page_config(page_title="Análise por Município", layout="wide") # Config é feito no app.py
st.sidebar.empty()
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

st.header("1. Configurações da Análise Municipal")
st.warning("⚠️ **Aviso de Performance:** Esta análise carrega arquivos de dados muito grandes (mais de 1.5 GB por ano) e **não funcionará** no plano gratuito do Streamlit Cloud (limite de 1GB RAM). Use uma plataforma com mais memória (como Hugging Face Spaces ou um servidor local).")

lista_de_municipios = obter_lista_de_municipios()
mapa_codigos_municipios = obter_mapa_codigos_municipios()
mapa_nomes_paises = obter_dados_paises()
ano_atual = datetime.now().year

col1, col2 = st.columns(2)
with col1:
    ano_principal = st.number_input(
        "Ano de Referência:", min_value=1998, max_value=ano_atual, value=ano_atual,
        help="O ano principal que você quer analisar."
    )
    municipios_selecionados = st.multiselect(
        "Selecione o(s) município(s):",
        options=lista_de_municipios,
        default=["Belo Horizonte"],
        help="Você pode digitar para pesquisar."
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

if st.button("Iniciar Análise por Município"):
    
    with st.spinner(f"Processando dados municipais para {', '.join(municipios_selecionados)}..."):
        try:
            # --- Validação ---
            codigos_municipios = [mapa_codigos_municipios.get(m) for m in municipios_selecionados if m in mapa_codigos_municipios]
            if not codigos_municipios:
                st.error("Nenhum município selecionado ou válido.")
                st.stop()
            
            # --- URLs ---
            url_exp_mun_principal = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/mun/EXP_{ano_principal}_MUN.csv"
            url_exp_mun_comparacao = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/mun/EXP_{ano_comparacao}_MUN.csv"
            url_imp_mun_principal = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/mun/IMP_{ano_principal}_MUN.csv"
            url_imp_mun_comparacao = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/mun/IMP_{ano_comparacao}_MUN.csv"

            # --- Carregamento ---
            df_exp_mun_princ = carregar_dataframe(url_exp_mun_principal, f"EXP_{ano_principal}_MUN.csv", usecols=MUN_COLS, dtypes=MUN_DTYPES)
            df_exp_mun_comp = carregar_dataframe(url_exp_mun_comparacao, f"EXP_{ano_comparacao}_MUN.csv", usecols=MUN_COLS, dtypes=MUN_DTYPES)
            df_imp_mun_princ = carregar_dataframe(url_imp_mun_principal, f"IMP_{ano_principal}_MUN.csv", usecols=MUN_COLS, dtypes=MUN_DTYPES)
            df_imp_mun_comp = carregar_dataframe(url_imp_mun_comparacao, f"IMP_{ano_comparacao}_MUN.csv", usecols=MUN_COLS, dtypes=MUN_DTYPES)

            if df_exp_mun_princ is None or df_imp_mun_princ is None or df_exp_mun_comp is None or df_imp_mun_comp is None:
                st.error("Falha ao carregar arquivos de dados municipais. Tente novamente.")
                st.stop()
            
            # --- Filtro de Meses ---
            if meses_selecionados:
                meses_para_filtrar = [MESES_MAPA[m] for m in meses_selecionados]
            else:
                meses_para_filtrar = list(range(1, df_exp_mun_princ['CO_MES'].max() + 1))

            # --- Processamento Exportação ---
            st.header(f"Exportações de {', '.join(municipios_selecionados)}")
            df_exp_mun_princ_f = df_exp_mun_princ[(df_exp_mun_princ['CO_MUN'].isin(codigos_municipios)) & (df_exp_mun_princ['CO_MES'].isin(meses_para_filtrar))]
            df_exp_mun_comp_f = df_exp_mun_comp[(df_exp_mun_comp['CO_MUN'].isin(codigos_municipios)) & (df_exp_mun_comp['CO_MES'].isin(meses_para_filtrar))]
            
            exp_paises_princ = df_exp_mun_princ_f.groupby('CO_PAIS')['VL_FOB'].sum().sort_values(ascending=False).reset_index()
            exp_paises_comp = df_exp_mun_comp_f.groupby('CO_PAIS')['VL_FOB'].sum().reset_index()
            
            # Mapeia nomes e formata valores
            exp_paises_princ['País'] = exp_paises_princ['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            exp_paises_princ[f'Valor {ano_principal} (US$)'] = exp_paises_princ['VL_FOB']
            
            exp_paises_comp['País'] = exp_paises_comp['CO_PAIS'].map(mapa_nomes_paises).fillna("Desconhecido")
            exp_paises_comp[f'Valor {ano_comparacao} (US$)'] = exp_paises_comp['VL_FOB']

            # Junta os dois anos
            exp_final = pd.merge(exp_paises_princ[['País', f'Valor {ano_principal} (US$)']], 
                                 exp_paises_comp[['País', f'Valor {ano_comparacao} (US$)']], 
                                 on="País", how="outer").fillna(0)
            
            # Calcula Variação
            exp_final['Variação %'] = 100 * (exp_final[f'Valor {ano_principal} (US$)'] - exp_final[f'Valor {ano_comparacao} (US$)']) / exp_final[f'Valor {ano_comparacao} (US$)']
            exp_final['Variação %'] = exp_final['Variação %'].fillna(0).round(2) # Substitui inf por 0

            # Formata valores em US$
            exp_final[f'Valor {ano_principal}'] = exp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
            exp_final[f'Valor {ano_comparacao}'] = exp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)
            
            st.dataframe(exp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).head(20)
                         [['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']])
            
            del df_exp_mun_princ, df_exp_mun_comp, df_exp_mun_princ_f, df_exp_mun_comp_f, exp_paises_princ, exp_paises_comp, exp_final

            # --- Processamento Importação ---
            st.header(f"Importações de {', '.join(municipios_selecionados)}")
            df_imp_mun_princ_f = df_imp_mun_princ[(df_imp_mun_princ['CO_MUN'].isin(codigos_municipios)) & (df_imp_mun_princ['CO_MES'].isin(meses_para_filtrar))]
            df_imp_mun_comp_f = df_imp_mun_comp[(df_imp_mun_comp['CO_MUN'].isin(codigos_municipios)) & (df_imp_mun_comp['CO_MES'].isin(meses_para_filtrar))]

            imp_paises_princ = df_imp_mun_princ_f.groupby('CO_PAIS')['VL_FOB'].sum().sort_values(ascending=False).reset_index()
            imp_paises_comp = df_imp_mun_comp_f.groupby('CO_PAIS')['VL_FOB'].sum().reset_index()

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

            del df_imp_mun_princ, df_imp_mun_comp, df_imp_mun_princ_f, df_imp_mun_comp_f, imp_paises_princ, imp_paises_comp, imp_final

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante a análise municipal:")
            st.exception(e)
