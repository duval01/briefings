import streamlit as st
import os

st.set_page_config(
    page_title="Gerador de Briefings | AEST",
    page_icon="📊",
    layout="wide"
)

st.title(" Central de Automações | AEST")
st.write("---")

# --- ALTERAÇÃO AQUI: Limpa a sidebar na home page ---
st.sidebar.empty() 
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)
# --- FIM DA ALTERAÇÃO ---

st.header("Bem-vindo à central de automações da AEST")
st.markdown("""
Esta é uma ferramenta automatizada para unificar as automações criadas pela AEST.

### 🧭 Como navegar

Use o menu lateral (à esquerda) para selecionar o tipo de análise que deseja realizar:

 1. Análise por País: Permite filtrar por um ou mais países e analisar o comércio de Minas Gerais com eles (produtos e municípios envolvidos).
 2. Análise por Município: Permite filtrar por um ou mais municípios de MG e analisar seus principais destinos/origens e produtos.
 3. Análise por Produto: Permite filtrar por NCM e analisar os principais destinos/origens e municípios.

""")

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





