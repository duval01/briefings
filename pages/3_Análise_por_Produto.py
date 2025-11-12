import streamlit as st
import pandas as pd
import requests
from io import StringIO
from urllib3.exceptions import InsecureRequestWarning
import os
from datetime import datetime
import io
import re
import zipfile
from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

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
        mapa_codigo_nome = pd.Series(df_pais.NO_PAIS.values, index=df_pais.CO_PAIS).to_dict()
        # --- NOVO: Retorna também a lista de nomes e o mapa reverso ---
        lista_nomes = sorted(df_pais['NO_PAIS'].unique().tolist())
        mapa_nome_codigo = pd.Series(df_pais.CO_PAIS.values, index=df_pais.NO_PAIS).to_dict()
        return mapa_codigo_nome, lista_nomes, mapa_nome_codigo
    return {}, [], {}

@st.cache_data
def obter_dados_produtos_ncm():
    """Carrega a tabela NCM completa (SH2, SH4 e SH6) e armazena em cache."""
    url_ncm = "https://balanca.economia.gov.br/balanca/bd/tabelas/NCM_SH.csv"
    # --- CORREÇÃO APLICADA: Usa CO_SH6 e NO_SH6_POR ---
    usecols_ncm = ['CO_SH2', 'NO_SH2_POR', 'CO_SH4', 'NO_SH4_POR', 'CO_SH6', 'NO_SH6_POR']
    df_ncm = carregar_dataframe(url_ncm, "NCM_SH.csv", usecols=usecols_ncm, mostrar_progresso=False)
    if df_ncm is not None:
        return df_ncm
    return None

def obter_lista_de_produtos_sh2():
    """Retorna uma lista de capítulos (SH2)."""
    df_ncm = obter_dados_produtos_ncm()
    if df_ncm is not None:
        df_sh2 = df_ncm.drop_duplicates(subset=['CO_SH2']).dropna()
        df_sh2['Display'] = df_sh2['CO_SH2'].astype(str).str.zfill(2) + " - " + df_sh2['NO_SH2_POR']
        lista_produtos = df_sh2['Display'].unique().tolist()
        lista_produtos.sort()
        return lista_produtos
    return ["Erro ao carregar lista de capítulos"]

def obter_lista_de_produtos_sh4(codigos_sh2_selecionados):
    """Retorna uma lista de produtos (SH4), opcionalmente filtrada por SH2."""
    df_ncm = obter_dados_produtos_ncm()
    if df_ncm is None:
        return ["Erro ao carregar lista de produtos"]

    df_sh4 = df_ncm.drop_duplicates(subset=['CO_SH4']).dropna(subset=['CO_SH4', 'NO_SH4_POR'])

    # Filtra por SH2 se algum for selecionado
    if codigos_sh2_selecionados:
        codigos_sh2_str = [str(c).zfill(2) for c in codigos_sh2_selecionados]
        df_sh4 = df_sh4[df_sh4['CO_SH2'].astype(str).str.zfill(2).isin(codigos_sh2_str)]

    df_sh4['Display'] = df_sh4['CO_SH4'].astype(str).str.zfill(4) + " - " + df_sh4['NO_SH4_POR']
    lista_produtos = df_sh4['Display'].unique().tolist()
    lista_produtos.sort()
    return lista_produtos

# --- NOVO: Função para obter lista de SH6 ---
@st.cache_data
def obter_lista_de_produtos_sh6(codigos_sh4_selecionados):
    """Retorna uma lista de produtos (SH6), opcionalmente filtrada por SH4."""
    df_ncm = obter_dados_produtos_ncm()
    if df_ncm is None:
        return ["Erro ao carregar lista de SH6"]

    # --- CORREÇÃO APLICADA: Usa CO_SH6 e NO_SH6_POR ---
    df_sh6 = df_ncm.drop_duplicates(subset=['CO_SH6']).dropna(subset=['CO_SH6', 'NO_SH6_POR'])

    # Filtra por SH4 se algum for selecionado
    if codigos_sh4_selecionados:
        codigos_sh4_str = [str(c).zfill(4) for c in codigos_sh4_selecionados]
        # Adiciona lógica para mapear SH4 -> SH6
        # CORREÇÃO AQUI:
        df_sh6['CO_SH4_TEMP'] = df_sh6['CO_SH6'].astype(str).str.zfill(6).str[:4]
        df_sh6 = df_sh6[df_sh6['CO_SH4_TEMP'].isin(codigos_sh4_str)]

    # --- CORREÇÃO APLICADA: Usa CO_SH6 e NO_SH6_POR ---
    df_sh6['Display'] = df_sh6['CO_SH6'].astype(str).str.zfill(6) + " - " + df_sh6['NO_SH6_POR']
    lista_produtos = df_sh6['Display'].unique().tolist()
    lista_produtos.sort()
    return lista_produtos


def get_sh4(co_ncm):
    """Extrai SH4 de um CO_NCM."""
    co_ncm_str = str(co_ncm).strip()
    if pd.isna(co_ncm_str) or co_ncm_str == "":
        return None
    co_ncm_str = co_ncm_str.zfill(8)
    return co_ncm_str[:4]

# --- NOVO: Função para extrair SH6 ---
def get_sh6(co_ncm):
    """Extrai SH6 de um CO_NCM."""
    co_ncm_str = str(co_ncm).strip()
    if pd.isna(co_ncm_str) or co_ncm_str == "":
        return None
    co_ncm_str = co_ncm_str.zfill(8)
    return co_ncm_str[:6] # Pega os 6 primeiros dígitos
# --- FIM NOVO ---

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

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def calcular_diferenca_percentual(valor_atual, valor_anterior):
    """Calcula a diferença percentual entre dois valores."""
    if valor_anterior == 0:
        return 0.0, "acréscimo" if valor_atual > 0 else "redução" if valor_atual < 0 else "estabilidade"
    diferenca = round(((valor_atual - valor_anterior) / valor_anterior) * 100, 2)
    if diferenca > 0:
        tipo_diferenca = "um acréscimo"
    elif diferenca < 0:
        tipo_diferenca = "uma redução"
    else:
        tipo_diferenca = "uma estabilidade"
    diferenca = abs(diferenca)
    return diferenca, tipo_diferenca

class DocumentoApp:
    def __init__(self, logo_path):
        self.doc = Document()
        self.secao_atual = 0
        self.subsecao_atual = 0
        self.titulo_doc = ""
        self.logo_path = logo_path
        self.diretorio_base = "/tmp/" 
    def set_titulo(self, titulo):
        self.titulo_doc = sanitize_filename(titulo)
        self.criar_cabecalho()
        p = self.doc.add_paragraph()
        run = p.add_run(self.titulo_doc)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    def adicionar_conteudo_formatado(self, texto):
        p = self.doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(1.25)
        run = p.add_run(texto)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    def adicionar_paragrafo(self, texto): 
        p = self.doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(1.25)
        run = p.add_run(texto)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    def adicionar_titulo(self, texto):
        p = self.doc.add_paragraph()
        if self.subsecao_atual == 0:
            run = p.add_run(f"{self.secao_atual}. {texto}")
        else:
            run = p.add_run(f"{self.secao_atual}.{self.subsecao_atual}. {texto}")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    def nova_secao(self):
        self.secao_atual += 1
        self.subsecao_atual = 0
    def criar_cabecalho(self):
        section = self.doc.sections[0]
        section.top_margin = Cm(1.27)
        header = section.header
        largura_total_cm = 16.0
        table = header.add_table(rows=1, cols=2, width=Cm(largura_total_cm))
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.columns[0].width = Cm(4.0)
        table.columns[1].width = Cm(12.0)
        cell_imagem = table.cell(0, 0)
        paragraph_imagem = cell_imagem.paragraphs[0]
        paragraph_imagem.paragraph_format.space_before = Pt(0)
        paragraph_imagem.paragraph_format.space_after = Pt(0)
        run_imagem = paragraph_imagem.add_run()
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                run_imagem.add_picture(self.logo_path,
                                       width=Cm(3.5), 
                                       height=Cm(3.42))
            except Exception as e:
                paragraph_imagem.add_run("[Logo não encontrado]")
        else:
            paragraph_imagem.add_run("[Logo não encontrado]")
        paragraph_imagem.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_texto = table.cell(0, 1)
        textos = [
            "GOVERNO DO ESTADO DE MINAS GERAIS",
            "SECRETARIA DE ESTADO DE DESENVOLVIMENTO ECONÔMICO",
            "Subsecretaria de Promoção de Investimentos e Cadeias Produtivas",
            "Superintendência de Atração de Investimentos e Estímulo à Exportação"
        ]
        def formatar_paragrafo_cabecalho(p):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(11)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p = cell_texto.paragraphs[0]
        formatar_paragrafo_cabecalho(p)
        run = p.add_run(textos[0])
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.bold = True 
        p = cell_texto.add_paragraph()
        formatar_paragrafo_cabecalho(p)
        run = p.add_run(textos[1])
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.bold = True
        for texto in textos[2:]: 
            p = cell_texto.add_paragraph()
            formatar_paragrafo_cabecalho(p)
            run = p.add_run(texto)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.bold = False 
    def finalizar_documento(self):
        diretorio_real = self.diretorio_base
        try:
            os.makedirs(diretorio_real, exist_ok=True)
        except Exception:
            diretorio_real = "/tmp/"
            os.makedirs(diretorio_real, exist_ok=True)
        nome_arquivo = f"{self.titulo_doc}.docx"
        nome_arquivo_sanitizado = sanitize_filename(nome_arquivo)
        file_stream = io.BytesIO()
        self.doc.save(file_stream)
        file_stream.seek(0)
        file_bytes = file_stream.getvalue()
        st.success(f"Documento '{nome_arquivo_sanitizado}' gerado com sucesso!")
        return file_bytes, nome_arquivo_sanitizado

# --- CONFIGURAÇÃO DA PÁGINA ---

st.header("1. Configurações da Análise de Produto")

# --- Callback para limpar o state ---
def clear_download_state_prod():
    """Limpa os relatórios gerados da sessão."""
    if 'arquivos_gerados_produto' in st.session_state:
        st.session_state.arquivos_gerados_produto = []

# --- ALTERADO: Carrega dados de Países aqui ---
lista_de_produtos_sh2 = obter_lista_de_produtos_sh2()
mapa_nomes_paises, lista_paises_nomes, mapa_paises_reverso = obter_dados_paises()
# --- FIM ALTERADO ---
ano_atual = datetime.now().year

col1, col2 = st.columns(2)
with col1:
    ano_principal = st.number_input(
        "Ano de Referência:", min_value=1998, max_value=ano_atual, value=ano_atual,
        help="O ano principal que você quer analisar.",
        on_change=clear_download_state_prod 
    )
    ano_comparacao = st.number_input(
        "Ano de Comparação:", min_value=1998, max_value=ano_atual, value=ano_atual - 1,
        help="O ano contra o qual você quer comparar.",
        on_change=clear_download_state_prod 
    )
    meses_selecionados = st.multiselect(
        "Meses de Análise (opcional):",
        options=LISTA_MESES,
        help="Selecione os meses. Se deixar em branco, o ano inteiro será analisado.",
        on_change=clear_download_state_prod 
    )

with col2:
    # --- NOVO: Filtro de País ---
    paises_selecionados_nomes = st.multiselect(
        "Filtrar por País (opcional):",
        options=lista_paises_nomes,
        help="Filtre a análise para países específicos (destino na EXP, origem na IMP).",
        on_change=clear_download_state_prod 
    )
    # --- FIM NOVO ---

    # --- Filtro SH2 ---
    sh2_selecionados_nomes = st.multiselect(
        "1. Filtrar por Capítulo (SH2) (opcional):",
        options=lista_de_produtos_sh2,
        help="Filtre a lista de produtos SH4 abaixo.",
        on_change=clear_download_state_prod 
    )
    codigos_sh2_selecionados = [int(s.split(" - ")[0]) for s in sh2_selecionados_nomes]
    
    # Lista de SH4 agora é filtrada pelo SH2 selecionado
    lista_de_produtos_sh4_filtrada = obter_lista_de_produtos_sh4(codigos_sh2_selecionados)
    
    sh4_selecionados_nomes = st.multiselect(
        "2. Selecione o(s) produto(s) (SH4) (opcional):",
        options=lista_de_produtos_sh4_filtrada, # Usa a lista filtrada
        default=[],
        help="Selecione um ou mais SH4. Isso filtrará a lista de SH6 abaixo.",
        on_change=clear_download_state_prod 
    )
    codigos_sh4_selecionados = [s.split(" - ")[0] for s in sh4_selecionados_nomes]

    # --- NOVO: Filtro SH6 ---
    lista_de_produtos_sh6_filtrada = obter_lista_de_produtos_sh6(codigos_sh4_selecionados)
    
    sh6_selecionados_nomes = st.multiselect(
        "3. Refinar por SH6 (opcional):",
        options=lista_de_produtos_sh6_filtrada,
        default=[],
        help="Selecione um ou mais SH6. Se selecionado, este filtro substitui o SH4.",
        on_change=clear_download_state_prod
    )
    # --- FIM NOVO ---


# --- Lógica de Agrupamento ---
agrupado = True
nome_agrupamento = None 
# --- ALTERADO: Lógica de agrupamento prioriza SH6 ---
if len(sh6_selecionados_nomes) > 1:
    nome_nivel = "SH6"
    produtos_para_agrupar_nomes = sh6_selecionados_nomes
elif len(sh4_selecionados_nomes) > 1 and not sh6_selecionados_nomes:
    nome_nivel = "SH4"
    produtos_para_agrupar_nomes = sh4_selecionados_nomes
else:
    produtos_para_agrupar_nomes = [] # Nenhum agrupamento

if len(produtos_para_agrupar_nomes) > 1:
    st.header("2. Opções de Agrupamento")
    agrupamento_input = st.radio(
        f"Deseja que os dados dos {len(produtos_para_agrupar_nomes)} produtos {nome_nivel} sejam agrupados?",
        ("agrupados", "separados"),
        index=0,
        horizontal=True,
        on_change=clear_download_state_prod 
    )
    agrupado = (agrupamento_input == "agrupados")
    
    if agrupado:
        quer_nome_agrupamento = st.checkbox(
            "Deseja dar um nome para este agrupamento de produtos?", 
            key="prod_nome_grupo",
            on_change=clear_download_state_prod
        )
        if quer_nome_agrupamento:
            nome_agrupamento = st.text_input(
                "Digite o nome do agrupamento:", 
                key="prod_nome_input",
                on_change=clear_download_state_prod
            )
    
    st.header("3. Gerar Análise")
else:
    agrupado = False # Se for só 1 produto (ou 0), a lógica é sempre "separado"
    st.header("2. Gerar Análise")
# --- FIM ALTERADO ---

# --- Inicialização do Session State ---
if 'arquivos_gerados_produto' not in st.session_state:
    st.session_state.arquivos_gerados_produto = []


if st.button("Iniciar Análise por Produto"):
    
    st.session_state.arquivos_gerados_produto = []
    logo_path_to_use = "LogoMinasGerais.png" 
    
    with st.spinner(f"Processando dados de produto..."):
        try:
            # --- Validação ---
            # --- ALTERADO: Prioriza SH6 ---
            codigos_sh6_selecionados = [s.split(" - ")[0] for s in sh6_selecionados_nomes]
            codigos_sh4_selecionados = [s.split(" - ")[0] for s in sh4_selecionados_nomes]
            
            if not codigos_sh6_selecionados and not codigos_sh4_selecionados:
                st.error("Nenhum produto (SH4 ou SH6) selecionado.")
                st.stop()
            
            # --- NOVO: Processa filtro de país ---
            codigos_paises_selecionados = [mapa_paises_reverso[nome] for nome in paises_selecionados_nomes]
            # --- FIM NOVO ---

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
                nome_periodo = f"o período de {', '.join(meses_selecionados)} de {ano_principal}"
                nome_periodo_comp = f"o mesmo período de {ano_comparacao}"
            else:
                meses_para_filtrar = list(range(1, df_exp_princ['CO_MES'].max() + 1))
                nome_periodo = f"o ano de {ano_principal} (completo)"
                nome_periodo_comp = f"o mesmo período de {ano_comparacao}"
            
            # --- Adiciona colunas SH4 e SH6 ---
            df_exp_princ['SH4'] = df_exp_princ['CO_NCM'].apply(get_sh4)
            df_exp_comp['SH4'] = df_exp_comp['CO_NCM'].apply(get_sh4)
            df_imp_princ['SH4'] = df_imp_princ['CO_NCM'].apply(get_sh4)
            df_imp_comp['SH4'] = df_imp_comp['CO_NCM'].apply(get_sh4)
            # --- NOVO ---
            df_exp_princ['SH6'] = df_exp_princ['CO_NCM'].apply(get_sh6)
            df_exp_comp['SH6'] = df_exp_comp['CO_NCM'].apply(get_sh6)
            df_imp_princ['SH6'] = df_imp_princ['CO_NCM'].apply(get_sh6)
            df_imp_comp['SH6'] = df_imp_comp['CO_NCM'].apply(get_sh6)
            # --- FIM NOVO ---

            # --- Lógica de Loop (Agrupado vs Separado) ---
            
            # --- ALTERADO: Define a lista de produtos a processar ---
            if agrupado:
                # Se agrupado, define o nome e a lista de códigos com base no nível (SH6 ou SH4)
                if codigos_sh6_selecionados:
                    codigos_loop = codigos_sh6_selecionados
                    nomes_base_loop = sh6_selecionados_nomes
                    nivel_filtro = "SH6"
                else:
                    codigos_loop = codigos_sh4_selecionados
                    nomes_base_loop = sh4_selecionados_nomes
                    nivel_filtro = "SH4"
                
                nome_grupo = nome_agrupamento if (nome_agrupamento and nome_agrupamento.strip() != "") else ", ".join([p.split(' - ')[1] for p in nomes_base_loop])
                # A lista de processamento tem 1 item: o grupo
                produtos_para_processar = [{"nome": nome_grupo, "codigos": codigos_loop, "nivel": nivel_filtro}]

            else:
                # Se separado, cria uma lista de dicionários para cada item
                produtos_para_processar = []
                if codigos_sh6_selecionados:
                    for nome_completo in sh6_selecionados_nomes:
                        produtos_para_processar.append({
                            "nome": nome_completo,
                            "codigos": [nome_completo.split(" - ")[0]], # Lista com 1 código
                            "nivel": "SH6"
                        })
                elif codigos_sh4_selecionados:
                     for nome_completo in sh4_selecionados_nomes:
                        produtos_para_processar.append({
                            "nome": nome_completo,
                            "codigos": [nome_completo.split(" - ")[0]], # Lista com 1 código
                            "nivel": "SH4"
                        })
            # --- FIM ALTERADO ---
            
            # Loop principal de processamento
            for produto_info in produtos_para_processar:
                
                app = DocumentoApp(logo_path=logo_path_to_use)
                
                # --- ALTERAÇÃO AQUI: Título dinâmico ---
                if agrupado:
                    st.subheader(f"Análise Agrupada de: {produto_info['nome']}")
                    nome_limpo_arquivo = sanitize_filename(produto_info['nome'])
                    titulo_doc = f"Briefing - {nome_limpo_arquivo} - {ano_principal}"
                    produto_nome_doc = f"de {produto_info['nome']}" # Nome para o texto
                else:
                    st.subheader(f"Análise de: {produto_info['nome']}")
                    nome_limpo_arquivo = sanitize_filename(produto_info['nome'].split(" - ")[1]) # Pega só o nome
                    titulo_doc = f"Briefing - {nome_limpo_arquivo} - {ano_principal}"
                    produto_nome_doc = f"de {produto_info['nome'].split(' - ')[1]}" # Nome para o texto
                
                app.set_titulo(titulo_doc)
                # --- FIM DA ALTERAÇÃO ---

                # --- Processamento Exportação ---
                st.header("Principais Destinos (Exportação de MG)")
                
                # --- ALTERADO: Lógica de filtragem ---
                # 1. Filtro base (MG e Mês)
                df_exp_princ_base = df_exp_princ[(df_exp_princ['SG_UF_NCM'] == 'MG') & (df_exp_princ['CO_MES'].isin(meses_para_filtrar))]
                df_exp_comp_base = df_exp_comp[(df_exp_comp['SG_UF_NCM'] == 'MG') & (df_exp_comp['CO_MES'].isin(meses_para_filtrar))]

                # 2. Filtro de Produto (SH6 ou SH4)
                nivel_filtro = produto_info['nivel'] # 'SH6' ou 'SH4'
                codigos_filtro = produto_info['codigos']
                df_exp_princ_f = df_exp_princ_base[df_exp_princ_base[nivel_filtro].isin(codigos_filtro)]
                df_exp_comp_f = df_exp_comp_base[df_exp_comp_base[nivel_filtro].isin(codigos_filtro)]
                
                # 3. NOVO: Filtro de País (se houver)
                if codigos_paises_selecionados:
                    df_exp_princ_f = df_exp_princ_f[df_exp_princ_f['CO_PAIS'].isin(codigos_paises_selecionados)]
                    df_exp_comp_f = df_exp_comp_f[df_exp_comp_f['CO_PAIS'].isin(codigos_paises_selecionados)]
                # --- FIM ALTERADO ---
                
                exp_total_princ = df_exp_princ_f['VL_FOB'].sum()
                exp_total_comp = df_exp_comp_f['VL_FOB'].sum()
                dif_exp, tipo_dif_exp = calcular_diferenca_percentual(exp_total_princ, exp_total_comp)
                
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
                exp_final['Variação %'] = exp_final['Variação %'].replace([float('inf'), float('-inf')], 0).fillna(0).round(2)

                exp_final[f'Valor {ano_principal}'] = exp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
                exp_final[f'Valor {ano_comparacao}'] = exp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)
                
                df_display_exp = exp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).reset_index(drop=True)
                st.dataframe(
                    df_display_exp[['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']].head(10),
                    hide_index=True
                )
                
                # --- Geração de Texto (Exportação) ---
                texto_exp_total = f"Em {nome_periodo}, as exportações de Minas Gerais {produto_nome_doc} somaram {formatar_valor(exp_total_princ)}, {tipo_dif_exp} de {dif_exp:.1f}% em relação a {nome_periodo_comp}."
                app.nova_secao()
                app.adicionar_titulo("Exportações de Produto")
                app.adicionar_conteudo_formatado(texto_exp_total)
                
                if exp_total_princ > 0:
                    texto_exp_paises = "Os principais países de destino foram: " + ", ".join(df_display_exp.head(5)['País'].tolist()) + "."
                    app.adicionar_conteudo_formatado(texto_exp_paises)
                
                del df_exp_princ_base, df_exp_comp_base, df_exp_princ_f, df_exp_comp_f, exp_paises_princ, exp_paises_comp, exp_final, df_display_exp

                # --- Processamento Importação ---
                st.header("Principais Origens (Importação de MG)")
                
                # --- ALTERADO: Lógica de filtragem ---
                # 1. Filtro base (MG e Mês)
                df_imp_princ_base = df_imp_princ[(df_imp_princ['SG_UF_NCM'] == 'MG') & (df_imp_princ['CO_MES'].isin(meses_para_filtrar))]
                df_imp_comp_base = df_imp_comp[(df_imp_comp['SG_UF_NCM'] == 'MG') & (df_imp_comp['CO_MES'].isin(meses_para_filtrar))]

                # 2. Filtro de Produto (SH6 ou SH4)
                df_imp_princ_f = df_imp_princ_base[df_imp_princ_base[nivel_filtro].isin(codigos_filtro)]
                df_imp_comp_f = df_imp_comp_base[df_imp_comp_base[nivel_filtro].isin(codigos_filtro)]
                
                # 3. NOVO: Filtro de País (se houver)
                if codigos_paises_selecionados:
                    df_imp_princ_f = df_imp_princ_f[df_imp_princ_f['CO_PAIS'].isin(codigos_paises_selecionados)]
                    df_imp_comp_f = df_imp_comp_f[df_imp_comp_f['CO_PAIS'].isin(codigos_paises_selecionados)]
                # --- FIM ALTERADO ---

                imp_total_princ = df_imp_princ_f['VL_FOB'].sum()
                imp_total_comp = df_imp_comp_f['VL_FOB'].sum()
                dif_imp, tipo_dif_imp = calcular_diferenca_percentual(imp_total_princ, imp_total_comp)

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
                imp_final['Variação %'] = imp_final['Variação %'].replace([float('inf'), float('-inf')], 0).fillna(0).round(2)
                
                imp_final[f'Valor {ano_principal}'] = imp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
                imp_final[f'Valor {ano_comparacao}'] = imp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)

                df_display_imp = imp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).reset_index(drop=True)
                st.dataframe(
                    df_display_imp[['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']].head(10),
                    hide_index=True
                )
                
                # --- Geração de Texto (Importação) ---
                texto_imp_total = f"Em {nome_periodo}, as importações de Minas Gerais {produto_nome_doc} somaram {formatar_valor(imp_total_princ)}, {tipo_dif_imp} de {dif_imp:.1f}% em relação a {nome_periodo_comp}."
                
                app.nova_secao()
                app.adicionar_titulo("Importações de Produto")
                app.adicionar_conteudo_formatado(texto_imp_total)
                
                if imp_total_princ > 0:
                    texto_imp_paises = "Os principais países de origem foram: " + ", ".join(df_display_imp.head(5)['País'].tolist()) + "."
                    app.adicionar_conteudo_formatado(texto_imp_paises)
                
                del df_imp_princ_base, df_imp_comp_base, df_imp_princ_f, df_imp_comp_f, imp_paises_princ, imp_paises_comp, imp_final, df_display_imp
            
                # Salva o documento no state
                file_bytes, file_name = app.finalizar_documento()
                st.session_state.arquivos_gerados_produto.append({"name": file_name, "data": file_bytes})
            
            # Limpa os DFs principais da memória após o loop
            del df_exp_princ, df_exp_comp, df_imp_princ, df_imp_comp

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante a análise de produto:")
            st.exception(e)

# --- Bloco de exibição de Download (COM LÓGICA DE ZIP) ---
if st.session_state.arquivos_gerados_produto:
    st.header("4. Relatórios Gerados")
    st.info("Clique para baixar os relatórios. Eles permanecerão aqui até que você gere um novo relatório.")
    
    if len(st.session_state.arquivos_gerados_produto) > 1:
        # Caso "Separados": Criar um ZIP
        st.subheader("Pacote de Relatórios (ZIP)")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for arquivo in st.session_state.arquivos_gerados_produto:
                zip_file.writestr(arquivo["name"], arquivo["data"])
        
        zip_bytes = zip_buffer.getvalue()
        
        st.download_button(
            label=f"Baixar todos os {len(st.session_state.arquivos_gerados_produto)} relatórios (.zip)",
            data=zip_bytes,
            file_name=f"Briefings_Produtos_{ano_principal}.zip",
            mime="application/zip",
            key="download_zip_produto"
        )
        
    elif len(st.session_state.arquivos_gerados_produto) == 1:
        # Caso "Agrupado": Botão único
        st.subheader("Relatório Gerado")
        arquivo = st.session_state.arquivos_gerados_produto[0] 
        st.download_button(
            label=f"Baixar Relatório ({arquivo['name']})",
            data=arquivo["data"], 
            file_name=arquivo["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"download_{arquivo['name']}"
        )

# --- Bloco de Rodapé ---
st.divider() 

col1, col2 = st.columns([0.3, 0.7], vertical_alignment="center") 

with col1:
    logo_footer_path = "AEST Sede.png"
    if os.path.exists(logo_footer_path):
        st.image(logo_footer_path, width=150)
    else:
        st.caption("Logo AEST não encontrada.")

with col2:
    st.caption("Desenvolvido por Aest - Dados e Subsecretaria de Promoção de Investimentos e Cadeias Produtivas")
