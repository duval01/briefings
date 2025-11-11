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

# --- CLASSE DOCUMENTO ---
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
        """Salva o documento em memória e retorna."""
        
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
# --- FIM DAS FUNÇÕES COPIADAS ---


# --- CONFIGURAÇÃO DA PÁGINA ---
st.sidebar.empty()
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

st.header("1. Configurações da Análise Municipal")
st.warning("⚠️ **Aviso de Performance:** Esta análise carrega arquivos de dados muito grandes (mais de 1.5 GB por ano) e **não funcionará** no plano gratuito do Streamlit Cloud (limite de 1GB RAM). Use uma plataforma com mais memória (como Hugging Face Spaces ou um servidor local).")

def clear_download_state_mun():
    """Limpa os relatórios gerados da sessão."""
    if 'arquivos_gerados_municipio' in st.session_state:
        st.session_state.arquivos_gerados_municipio = []

lista_de_municipios = obter_lista_de_municipios()
mapa_codigos_municipios = obter_mapa_codigos_municipios()
mapa_nomes_paises = obter_dados_paises()
ano_atual = datetime.now().year

col1, col2 = st.columns(2)
with col1:
    ano_principal = st.number_input(
        "Ano de Referência:", min_value=1998, max_value=ano_atual, value=ano_atual,
        help="O ano principal que você quer analisar.",
        on_change=clear_download_state_mun
    )
    municipios_selecionados = st.multiselect(
        "Selecione o(s) município(s):",
        options=lista_de_municipios,
        default=["BELO HORIZONTE"],
        help="Você pode digitar para pesquisar.",
        on_change=clear_download_state_mun
    )

with col2:
    ano_comparacao = st.number_input(
        "Ano de Comparação:", min_value=1998, max_value=ano_atual, value=ano_atual - 1,
        help="O ano contra o qual você quer comparar.",
        on_change=clear_download_state_mun
    )
    meses_selecionados = st.multiselect(
        "Meses de Análise (opcional):",
        options=LISTA_MESES,
        help="Selecione os meses. Se deixar em branco, o ano inteiro será analisado.",
        on_change=clear_download_state_mun
    )

# --- ALTERAÇÃO AQUI: Lógica de Agrupamento com Nome ---
agrupado = True
nome_agrupamento = None
if len(municipios_selecionados) > 1:
    st.header("2. Opções de Agrupamento")
    agrupamento_input = st.radio(
        "Deseja que os dados sejam agrupados ou separados?",
        ("agrupados", "separados"),
        index=0,
        horizontal=True,
        on_change=clear_download_state_mun
    )
    agrupado = (agrupamento_input == "agrupados")
    
    if agrupado:
        quer_nome_agrupamento = st.checkbox(
            "Deseja dar um nome para este agrupamento de municípios?", 
            key="mun_nome_grupo",
            on_change=clear_download_state_mun
        )
        if quer_nome_agrupamento:
            nome_agrupamento = st.text_input(
                "Digite o nome do agrupamento:", 
                key="mun_nome_input",
                on_change=clear_download_state_mun
            )

    st.header("3. Gerar Análise")
else:
    st.header("2. Gerar Análise")
# --- FIM DA ALTERAÇÃO ---


# --- Inicialização do Session State ---
if 'arquivos_gerados_municipio' not in st.session_state:
    st.session_state.arquivos_gerados_municipio = []


if st.button("Iniciar Análise por Município"):
    
    st.session_state.arquivos_gerados_municipio = []
    logo_path_to_use = "LogoMinasGerais.png"
    
    with st.spinner(f"Processando dados municipais para {', '.join(municipios_selecionados)}..."):
        try:
            # --- Validação ---
            codigos_municipios_map = [mapa_codigos_municipios.get(m) for m in municipios_selecionados if m in mapa_codigos_municipios]
            if not codigos_municipios_map:
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
                nome_periodo = f"o período de {', '.join(meses_selecionados)} de {ano_principal}"
                nome_periodo_comp = f"o mesmo período de {ano_comparacao}"
            else:
                meses_para_filtrar = list(range(1, df_exp_mun_princ['CO_MES'].max() + 1))
                nome_periodo = f"o ano de {ano_principal} (completo)"
                nome_periodo_comp = f"o mesmo período de {ano_comparacao}"

            
            # --- Lógica de Loop (Agrupado vs Separado) ---
            if not agrupado:
                municipios_para_processar = municipios_selecionados
            else:
                municipios_para_processar = [nome_agrupamento if nome_agrupamento else ", ".join(municipios_selecionados)]

            for municipio_nome in municipios_para_processar:
                
                app = DocumentoApp(logo_path=logo_path_to_use)
                
                if agrupado:
                    st.subheader(f"Análise Agrupada de: {municipio_nome}")
                    codigos_municipios_loop = codigos_municipios_map
                    nome_limpo_arquivo = sanitize_filename(municipio_nome)
                    titulo_doc = f"Briefing - {nome_limpo_arquivo} - {ano_principal}"
                    nome_doc = f"de {municipio_nome}"
                else:
                    st.subheader(f"Análise de: {municipio_nome}")
                    codigos_municipios_loop = [mapa_codigos_municipios.get(municipio_nome)]
                    nome_limpo_arquivo = sanitize_filename(municipio_nome)
                    titulo_doc = f"Briefing - {nome_limpo_arquivo} - {ano_principal}"
                    nome_doc = f"de {municipio_nome}"
                
                app.set_titulo(titulo_doc)

                # --- Processamento Exportação ---
                st.header("Principais Destinos (Exportação)")
                df_exp_mun_princ_f = df_exp_mun_princ[(df_exp_mun_princ['CO_MUN'].isin(codigos_municipios_loop)) & (df_exp_mun_princ['CO_MES'].isin(meses_para_filtrar))]
                df_exp_mun_comp_f = df_exp_mun_comp[(df_exp_mun_comp['CO_MUN'].isin(codigos_municipios_loop)) & (df_exp_mun_comp['CO_MES'].isin(meses_para_filtrar))]
                
                exp_total_princ = df_exp_mun_princ_f['VL_FOB'].sum()
                exp_total_comp = df_exp_mun_comp_f['VL_FOB'].sum()
                dif_exp, tipo_dif_exp = calcular_diferenca_percentual(exp_total_princ, exp_total_comp)
                
                exp_paises_princ = df_exp_mun_princ_f.groupby('CO_PAIS')['VL_FOB'].sum().sort_values(ascending=False).reset_index()
                exp_paises_comp = df_exp_mun_comp_f.groupby('CO_PAIS')['VL_FOB'].sum().reset_index()
                
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
                
                texto_exp_total = f"Em {nome_periodo}, as exportações {nome_doc} somaram {formatar_valor(exp_total_princ)}, {tipo_dif_exp} de {dif_exp:.1f}% em relação a {nome_periodo_comp}."
                
                app.nova_secao()
                app.adicionar_titulo("Exportações do Município")
                app.adicionar_conteudo_formatado(texto_exp_total)
                
                if exp_total_princ > 0:
                    texto_exp_paises = "Os principais países de destino foram: " + ", ".join(df_display_exp.head(5)['País'].tolist()) + "."
                    app.adicionar_conteudo_formatado(texto_exp_paises)
                
                del df_exp_mun_princ_f, df_exp_mun_comp_f, exp_paises_princ, exp_paises_comp, exp_final, df_display_exp

                # --- Processamento Importação ---
                st.header(f"Importações de {municipio_nome}")
                df_imp_mun_princ_f = df_imp_mun_princ[(df_imp_mun_princ['CO_MUN'].isin(codigos_municipios_loop)) & (df_imp_mun_princ['CO_MES'].isin(meses_para_filtrar))]
                df_imp_mun_comp_f = df_imp_mun_comp[(df_imp_mun_comp['CO_MUN'].isin(codigos_municipios_loop)) & (df_imp_mun_comp['CO_MES'].isin(meses_para_filtrar))]

                imp_total_princ = df_imp_mun_princ_f['VL_FOB'].sum()
                imp_total_comp = df_imp_mun_comp_f['VL_FOB'].sum()
                dif_imp, tipo_dif_imp = calcular_diferenca_percentual(imp_total_princ, imp_total_comp)

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
                imp_final['Variação %'] = imp_final['Variação %'].replace([float('inf'), float('-inf')], 0).fillna(0).round(2)
                
                imp_final[f'Valor {ano_principal}'] = imp_final[f'Valor {ano_principal} (US$)'].apply(formatar_valor)
                imp_final[f'Valor {ano_comparacao}'] = imp_final[f'Valor {ano_comparacao} (US$)'].apply(formatar_valor)

                df_display_imp = imp_final.sort_values(by=f'Valor {ano_principal} (US$)', ascending=False).reset_index(drop=True)
                st.dataframe(
                    df_display_imp[['País', f'Valor {ano_principal}', f'Valor {ano_comparacao}', 'Variação %']].head(10),
                    hide_index=True
                )
                
                texto_imp_total = f"Em {nome_periodo}, as importações {nome_doc} somaram {formatar_valor(imp_total_princ)}, {tipo_dif_imp} de {dif_imp:.1f}% em relação a {nome_periodo_comp}."
                
                app.nova_secao()
                app.adicionar_titulo("Importações do Município")
                app.adicionar_conteudo_formatado(texto_imp_total)

                if imp_total_princ > 0:
                    texto_imp_paises = "Os principais países de origem foram: " + ", ".join(df_display_imp.head(5)['País'].tolist()) + "."
                    app.adicionar_conteudo_formatado(texto_imp_paises)
                
                del df_imp_mun_princ_f, df_imp_mun_comp_f, imp_paises_princ, imp_paises_comp, imp_final, df_display_imp

                # Salva o documento no state
                file_bytes, file_name = app.finalizar_documento()
                st.session_state.arquivos_gerados_municipio.append({"name": file_name, "data": file_bytes})
            
            # Limpa os DFs principais da memória
            del df_exp_mun_princ, df_exp_mun_comp, df_imp_mun_princ, df_imp_mun_comp

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante a análise municipal:")
            st.exception(e)

# --- Bloco de Download (com ZIP) ---
if st.session_state.arquivos_gerados_municipio:
    st.header("4. Relatórios Gerados")
    st.info("Clique para baixar os relatórios. Eles permanecerão aqui até que você gere um novo relatório.")

    
    if len(st.session_state.arquivos_gerados_municipio) > 1:
        # ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for arquivo in st.session_state.arquivos_gerados_municipio:
                zip_file.writestr(arquivo["name"], arquivo["data"])
        
        zip_bytes = zip_buffer.getvalue()
        
        st.download_button(
            label=f"Baixar todos os {len(st.session_state.arquivos_gerados_municipio)} relatórios (.zip)",
            data=zip_bytes,
            file_name=f"Briefings_Municipios_{ano_principal}.zip",
            mime="application/zip",
            key="download_zip_municipio"
        )
        
    elif len(st.session_state.arquivos_gerados_municipio) == 1:
        # Botão único
        arquivo = st.session_state.arquivos_gerados_municipio[0] 
        st.download_button(
            label=f"Baixar Relatório ({arquivo['name']})",
            data=arquivo["data"], 
            file_name=arquivo["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"download_{arquivo['name']}"
        )
