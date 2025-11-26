import streamlit as st
import os
import time
# --- ALTERAÇÃO: Importa do novo auth.py ---
from auth import check_login, draw_sidebar, logout

# 1. Configuração da página (DEVE ser o primeiro comando)
st.set_page_config(
    page_title="Central de automações | AEST",
    page_icon="📊",
    layout="wide"
)

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
# -------------------------------------

# --- INICIALIZAÇÃO DO SESSION STATE ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_name = ""
    st.session_state.role = "guest"
    st.session_state.allowed_pages = {}

# --- LÓGICA PRINCIPAL DA PÁGINA ---
st.session_state.current_page = 'Home'
draw_sidebar()

if not st.session_state.logged_in:
    st.title(" Central de automações | AEST")
    st.write("---")
    st.header("Login")
    
    username = st.text_input("Usuário", key="login_user")
    password = st.text_input("Senha", type="password", key="login_pass")
    
    if st.button("Entrar", type="primary"):
        if check_login(username, password):
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")
else:
    st.title(" Central de automações | AEST")
    st.write("---")

    st.header(f"Bem-vindo(a) à central de automações, {st.session_state.user_name}!")
    st.markdown("""
    Esta é uma ferramenta criada para centralizar todas as automações da AEST.

    ### 🧭 Como navegar

    Use o menu de navegação (à esquerda) para selecionar a automação que deseja executar:
    """)

    allowed_pages = st.session_state.get('allowed_pages', {})
    if not allowed_pages:
        st.warning("Seu usuário não tem permissão para acessar nenhuma automação no momento.")
    else:
        for page_name, page_info in allowed_pages.items():
            if st.button(f"Acessar {page_name} {page_info.get('icon', '')}"):
                st.switch_page(page_info["path"])

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