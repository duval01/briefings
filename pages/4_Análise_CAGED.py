import streamlit as st
import pandas as pd
import os
import ftplib 
import py7zr 
from io import StringIO
from datetime import datetime
import time
import glob
import dask.dataframe as dd 

# --- OCULTA A NAVEGAÇÃO PADRÃO ---
st.markdown(
    """
    <style>
        [data-testid="stSidebarNav"] {
            display: none;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# --- IMPORTAÇÃO E PROTEÇÃO DA PÁGINA ---
try:
    from Home import draw_sidebar, logout
except ImportError:
    def draw_sidebar():
        st.sidebar.error("Erro ao carregar a navegação. Execute a partir do Home.py.")
    def logout():
        st.sidebar.error("Erro ao carregar.")

st.session_state.current_page = 'Análise CAGED' 
draw_sidebar()

if not st.session_state.get('logged_in', False):
    st.error("Acesso negado. Por favor, faça o login na Página Principal.")
    st.page_link("Home.py", label="Ir para a página de Login", icon="🏠")
    st.stop()
    
if st.session_state.get('role') != 'admin':
    st.error("Acesso negado. Você não tem permissão para ver esta página.")
    st.page_link("Home.py", label="Voltar à Página Principal", icon="🏠")
    st.stop()


# --- FUNÇÕES DE LÓGICA ---

DTYPES_BASE = {
    'município': 'str', 'seção': 'str', 'subclasse': 'str',
    'cbo2002ocupação': 'str', 'categoria': 'str', 'graudeinstrução': 'str',
    'raçacor': 'str', 'sexo': 'str', 'tipoempregador': 'str',
    'tipoestabelecimento': 'str', 'tipomovimentação': 'str',
    'tipodedeficiência': 'str', 'unidadesaláriocódigo': 'str',
    'idade': 'str', 'horascontratuais': 'str',
    'competênciamov': 'str', 'região': 'str', 'uf': 'str',
    'saldomovimentação': 'str', 'indtrabintermitente': 'str',
    'indtrabparcial': 'str', 'tamestabjan': 'str', 'indicadoraprendiz': 'str',
    'origemdainformação': 'str', 'competênciadec': 'str',
    'indicadordeforadoprazo': 'str'
}

DTYPES_MAP = {
    "Movimentações": DTYPES_BASE,
    "Fora de prazo": DTYPES_BASE,
    "Exclusões": {
        **DTYPES_BASE,
        'competênciaexc': 'str',
        'indicadordeexclusão': 'str'
    }
}

def baixar_e_processar_caged(tipo_caged, ano, mes_inicial, mes_final, pasta_local, tipos_de_arquivo_filtrados):
    """
    Processa dados do CAGED com base no tipo e na lista de arquivos selecionados.
    """
    try:
        os.makedirs(pasta_local, exist_ok=True)
        st.info(f"Pasta de trabalho: '{pasta_local}'")

        if tipo_caged in ["NOVO CAGED", "CAGED_AJUSTES"]:
            meses_para_processar = range(mes_inicial, mes_final + 1)
            estrutura_mensal = True
        else: 
            meses_para_processar = [0] 
            estrutura_mensal = False

        for mes in meses_para_processar:
            
            log_placeholder = st.empty()
            log_placeholder.info("Conectando ao servidor FTP...")
            
            try:
                ftp = ftplib.FTP('ftp.mtps.gov.br', timeout=60)
                ftp.login()
                log_placeholder.success("Conexão bem-sucedida.")
                
                if estrutura_mensal:
                    st.write("---")
                    st.subheader(f"Processando Mês/Ano: {mes:02d}/{ano}")
                    caminho_ftp = f'/pdet/microdados/{tipo_caged}/{ano}/{ano}{mes:02d}'
                else: 
                    st.write("---")
                    st.subheader(f"Processando Ano: {ano} (CAGED Antigo)")
                    caminho_ftp = f'/pdet/microdados/CAGED/{ano}'

                log_placeholder.info(f"Navegando para: {caminho_ftp}")
                ftp.cwd(caminho_ftp)
                arquivos_na_pasta = ftp.nlst()
                
            except ftplib.error_perm as e:
                # Loga o erro mas continua
                msg_erro = f"ERRO: A pasta não foi encontrada ({e})."
                st.error(msg_erro)
                continue 

            # Loop pelos tipos FILTRADOS
            for prefixo, nome_amigavel in tipos_de_arquivo_filtrados.items():
                
                st.write(f"**Tipo: {nome_amigavel.upper()}**")
                nome_arquivo_7z = None
                
                for arquivo in arquivos_na_pasta:
                    if not estrutura_mensal:
                        if arquivo.startswith(prefixo) and str(ano) in arquivo:
                            nome_arquivo_7z = arquivo
                            break
                    else:
                        if arquivo.startswith(prefixo):
                            nome_arquivo_7z = arquivo
                            break
                
                if not nome_arquivo_7z:
                    st.warning(f"AVISO: Arquivo com prefixo '{prefixo}' não encontrado. Pulando.")
                    continue

                caminho_local_7z = os.path.join(pasta_local, nome_arquivo_7z)

                log_placeholder.info(f"Baixando '{nome_arquivo_7z}'...")
                with open(caminho_local_7z, 'wb') as f:
                    ftp.retrbinary(f'RETR {nome_arquivo_7z}', f.write)
                log_placeholder.success("Download concluído.")

                pasta_sufixo = f"{prefixo}_{ano}_{mes:02d}" if estrutura_mensal else f"{prefixo}_{ano}"
                pasta_extracao = os.path.join(pasta_local, f"extracao_{pasta_sufixo}")
                
                log_placeholder.info(f"Descompactando para '{pasta_extracao}'...")
                with py7zr.SevenZipFile(caminho_local_7z, mode='r') as z:
                    z.extractall(path=pasta_extracao)

                arquivo_txt_nome = next((f for f in os.listdir(pasta_extracao) if f.lower().endswith('.txt')), None)
                
                if not arquivo_txt_nome:
                    st.error("ERRO: Nenhum arquivo .txt encontrado na pasta extraída.")
                    continue

                caminho_arquivo_txt = os.path.join(pasta_extracao, arquivo_txt_nome)

                log_placeholder.info("Lendo microdados...")
                try:
                    # Usa encoding latin-1 para ler os originais
                    dataframe = pd.read_csv(caminho_arquivo_txt, sep=';', encoding='latin-1', decimal=',', low_memory=False)
                    log_placeholder.success(f"{len(dataframe)} linhas carregadas.")
                except Exception as read_e:
                    st.error(f"Falha ao ler o arquivo: {read_e}")
                    continue

                nome_sufixo = f"{nome_amigavel}_{ano}_{mes:02d}" if estrutura_mensal else f"{nome_amigavel}_{ano}"
                nome_arquivo_final = f"caged_{nome_sufixo}.csv"
                caminho_final_csv = os.path.join(pasta_local, nome_arquivo_final)
                
                log_placeholder.info(f"Salvando arquivo final: '{caminho_final_csv}'")
                # Salva como UTF-8
                dataframe.to_csv(caminho_final_csv, index=False, encoding='utf-8-sig', sep=';')
                
                os.remove(caminho_local_7z)
                for f in os.listdir(pasta_extracao):
                    os.remove(os.path.join(pasta_extracao, f))
                os.rmdir(pasta_extracao)
                log_placeholder.success("Arquivos temporários limpos.")

            ftp.quit()

        st.write("---")
        st.success(f"=== PROCESSAMENTO DO ANO {ano} CONCLUÍDO! ===")

    except Exception as e:
        st.error(f"Ocorreu um erro geral: {e}")
        try: ftp.quit()
        except: pass


def concatenar_com_dask(pasta_dos_arquivos, padrao_glob, arquivo_final_csv, nome_amigavel, log_placeholder):
    try:
        log_placeholder.info(f"Iniciando concatenação para: {arquivo_final_csv}...")
        inicio = time.time()

        tipos_de_dados = DTYPES_MAP.get(nome_amigavel)
        if tipos_de_dados is None:
            tipos_de_dados = DTYPES_MAP['movimentacoes'] 

        caminho_busca = os.path.join(pasta_dos_arquivos, padrao_glob)
        arquivos_para_concatenar = glob.glob(caminho_busca)

        if not arquivos_para_concatenar:
            log_placeholder.warning(f"Nenhum arquivo encontrado com o padrão '{padrao_glob}'. Concatenação pulada.")
            return

        log_placeholder.info(f"Encontrados {len(arquivos_para_concatenar)} arquivos. Lendo metadados...")

        ddf = dd.read_csv(
            arquivos_para_concatenar, 
            sep=';',
            encoding='utf-8', 
            decimal=',', 
            dtype=tipos_de_dados,
            low_memory=False
        )
        
        log_placeholder.info(f"Metadados lidos. Partições: {ddf.npartitions}. Computando e salvando...")

        caminho_final = os.path.join(pasta_dos_arquivos, arquivo_final_csv)
        ddf.to_csv(caminho_final, single_file=True, index=False, encoding='utf-8-sig', sep=';')
        
        fim = time.time()
        tempo_total = fim - inicio

        st.success(f"Arquivo concatenado salvo em: {caminho_final} (Tempo: {tempo_total:.2f}s)")

    except Exception as e:
        st.error(f"Ocorreu um erro durante a concatenação Dask para '{nome_amigavel}': {e}")


# --- Interface ---

st.title("Automação de Microdados do CAGED")
st.info("Esta automação baixa, descompacta e processa os microdados do CAGED e os salva como arquivos .csv em uma pasta local.")

# --- CONFIGURAÇÕES ---
st.header("1. Configurações")
ano_atual = datetime.now().year

tipo_caged_selecionado = st.selectbox(
    "Selecione o tipo de microdado CAGED:",
    options=["NOVO CAGED", "CAGED_AJUSTES", "CAGED (Antigo)"],
    index=0,
    help="Selecione a pasta de dados no FTP."
)

if tipo_caged_selecionado in ["NOVO CAGED", "CAGED_AJUSTES"]:
    lista_anos_disponiveis = list(range(2020, ano_atual + 1))
    default_ano = [2024] if 2024 in lista_anos_disponiveis else [lista_anos_disponiveis[-1]]
    
    anos_selecionados = st.multiselect("Selecione o(s) Ano(s):", options=lista_anos_disponiveis, default=default_ano)
    col1, col2 = st.columns(2)
    with col1:
        mes_inicial = st.number_input("Mês Inicial:", min_value=1, max_value=12, value=1)
    with col2:
        mes_final = st.number_input("Mês Final:", min_value=1, max_value=12, value=12)
else: 
    lista_anos_disponiveis = list(range(2007, 2020))
    anos_selecionados = st.multiselect("Selecione o(s) Ano(s):", options=lista_anos_disponiveis, default=[2019])
    mes_inicial = 0
    mes_final = 0

pasta_local = st.text_input("Pasta local para salvar os arquivos:", 
                            r"C:\temp\caged_processado",
                            help="Caminho completo da pasta no seu computador.")

# --- NOVO: Seletor de Tipos de Arquivo ---
tipos_de_arquivo_base = {
    'CAGEDMOV': 'Movimentações',
    'CAGEDEXC': 'Exclusões',
    'CAGEDFOR': 'Fora de prazo'
}

opcoes_filtro = ["Todos"] + list(tipos_de_arquivo_base.values())

tipos_selecionados_ui = st.multiselect(
    "Quais tipos de arquivo você deseja baixar/processar?",
    options=opcoes_filtro,
    default=["Todos"],
    help="Selecione 'Todos' para baixar os 3 tipos, ou selecione tipos específicos."
)

# Lógica para filtrar o dicionário base
if "Todos" in tipos_selecionados_ui or not tipos_selecionados_ui:
    tipos_de_arquivo_final = tipos_de_arquivo_base
else:
    # Cria um novo dicionário apenas com os itens selecionados
    # Inverte o dicionário base para buscar pela chave (valor)
    tipos_de_arquivo_final = {}
    for k, v in tipos_de_arquivo_base.items():
        if v in tipos_selecionados_ui:
            tipos_de_arquivo_final[k] = v
# --- FIM DO NOVO SELETOR ---

concatenar_selecionado = st.checkbox(
    "Concatenar arquivos baixados por tipo?", 
    value=True,
    help="Junta todos os arquivos de um mesmo tipo em um único CSV grande."
)

# --- Botão de Execução ---
st.header("2. Executar Automação")
if st.button("Iniciar Download e Processamento", type="primary"):
    if not pasta_local:
        st.error("Por favor, defina uma pasta local.")
    elif not os.path.isdir(os.path.dirname(pasta_local)):
        st.error(f"Erro: O caminho-pai '{os.path.dirname(pasta_local)}' não existe.")
    elif not anos_selecionados:
        st.error("Selecione pelo menos um ano.")
    elif not tipos_de_arquivo_final:
        st.error("Selecione pelo menos um tipo de arquivo.")
    else:
        with st.spinner("Executando automação em lote..."):
            
            # 1. Download com tipos filtrados
            for ano in sorted(anos_selecionados):
                st.title(f"Processando Ano: {ano} ({tipo_caged_selecionado})")
                # Passa o dicionário filtrado
                baixar_e_processar_caged(tipo_caged_selecionado, ano, mes_inicial, mes_final, pasta_local, tipos_de_arquivo_final)
            
            st.title("Download concluído!")

            # 2. Concatenação com tipos filtrados
            if concatenar_selecionado:
                st.header("3. Concatenação de Arquivos")
                concat_log_placeholder = st.empty()
                
                if tipo_caged_selecionado == "CAGED (Antigo)":
                    tipos_para_concatenar = ['movimentacoes']
                else:
                    # Usa apenas os valores filtrados
                    tipos_para_concatenar = tipos_de_arquivo_final.values()
                
                anos_str = "-".join(map(str, sorted(anos_selecionados)))
                
                for nome_amigavel in tipos_para_concatenar:
                    padrao_glob = f"caged_{nome_amigavel}_*.csv"
                    nome_final = f"{tipo_caged_selecionado}_{nome_amigavel}_{anos_str}_COMPLETO.csv"
                    
                    concatenar_com_dask(pasta_local, padrao_glob, nome_final, nome_amigavel, concat_log_placeholder)
                
                st.success("Processo de concatenação finalizado!")