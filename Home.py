import streamlit as st
import os

# 1. Configuração da página (deve ser o primeiro comando)
st.set_page_config(
    page_title="Briefings ComexStat",
    page_icon="📊",
    layout="wide"
)

# 2. Logo da Sidebar (colocada aqui, ela fica ACIMA da navegação e em TODAS as páginas)
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

# 3. Conteúdo da Página Home
st.title(" automação de Briefings ComexStat")
st.write("---")

st.header("Bem-vindo ao Gerador de Briefings!")
st.markdown("""
Esta é uma ferramenta automatizada para criar relatórios de comércio exterior com base nos microdados públicos do ComexStat.

### 🧭 Como Navegar

Use o menu de navegação (à esquerda) para selecionar o tipo de análise que deseja realizar:

* **Análise por País:** Permite filtrar por um ou mais países e analisar o comércio de Minas Gerais com eles (incluindo os principais produtos e municípios envolvidos).
* **Análise por Município:** Permite filtrar por um ou mais municípios de MG e analisar seus principais destinos/origens (países).
* **Análise por Produto:** Permite filtrar por NCM e analisar os principais destinos/origens (países).

""")

# --- Bloco de Rodapé ---
st.divider() 

col1, col2 = st.columns([0.3, 0.7], vertical_alignment="center") 

with col1:
    # Coluna 1 (menor) agora contém a logo
    logo_footer_path = "AEST Sede.png"
    if os.path.exists(logo_footer_path):
        st.image(logo_footer_path, width=150)
    else:
        st.caption("Logo AEST não encontrada.")

with col2:
    # Coluna 2 (maior) agora contém o texto
    st.caption("Desenvolvido por Aest - Dados e Subsecretaria de Promoção de Investimentos e Cadeias Produtivas")
