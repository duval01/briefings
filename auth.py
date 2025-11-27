import streamlit as st
import os

# --- BANCO DE DADOS DE USUÁRIOS E PERMISSÕES ---
USERS = {
    "AEST": {
        "password": "aest123", 
        "role": "admin",
        "name": "Admin AEST"
    },
    "DIPEX": {
        "password": "dipex123", 
        "role": "dpex",
        "name": "Usuário DPEX"
    }
}

PAGES_CONFIG = {
    "admin": {
        "Briefings de País": {"path": "pages/1_Análise_por_País.py", "icon": "🌎"},
        "Briefings de Município": {"path": "pages/2_Análise_por_Município.py", "icon": "🏙️"},
        "Briefings de Produto": {"path": "pages/3_Análise_por_Produto.py", "icon": "📦"},
        "Automação CAGED": {"path": "pages/4_Análise_CAGED.py", "icon": "📄"}
    },
    "dpex": {
        "Briefings de País": {"path": "pages/1_Análise_por_País.py", "icon": "🌎"},
        "Briefings de Município": {"path": "pages/2_Análise_por_Município.py", "icon": "🏙️"},
        "Briefings de Produto": {"path": "pages/3_Análise_por_Produto.py", "icon": "📦"},
    }
}

def check_login(username, password):
    if username in USERS and USERS[username]["password"] == password:
        st.session_state.logged_in = True
        st.session_state.user_name = USERS[username]["name"]
        st.session_state.role = USERS[username]["role"]
        st.session_state.allowed_pages = PAGES_CONFIG.get(st.session_state.role, {})
        return True
    return False

def logout():
    st.session_state.logged_in = False
    st.session_state.user_name = ""
    st.session_state.role = "guest"
    st.session_state.allowed_pages = {}
    st.rerun()

def draw_sidebar():
    with st.sidebar:
        logo_sidebar_path = "LogoMinasGerais.png"
        if os.path.exists(logo_sidebar_path):
            st.image(logo_sidebar_path, width=200)
        
        if st.session_state.get('logged_in', False):
            st.write(f"Bem-vindo(a), **{st.session_state.user_name}**!")
            st.divider()
            
            allowed_pages = st.session_state.get('allowed_pages', {})
            
            st.page_link("Home.py", label="Página Principal", icon="🏠")
            
            for page_name, page_info in allowed_pages.items():
                st.page_link(page_info["path"], label=page_name, icon=page_info["icon"])
            
            st.divider()
            # --- CORREÇÃO DO ERRO AQUI (Adicionado key unique) ---
            if st.button("Sair (Logout)", key="logout_sidebar_btn"):
                logout()
        else:
            if st.session_state.get('current_page', 'Home') != 'Home':
                 st.page_link("Home.py", label="Ir para Login", icon="🏠")
            st.info("Por favor, faça o login para acessar as automações.")

def page_protector(page_name="Página", required_role=None):
    st.markdown("""<style>[data-testid="stSidebarNav"] {display: none;}</style>""", unsafe_allow_html=True)
    st.session_state.current_page = page_name
    draw_sidebar()

    if not st.session_state.get('logged_in', False):
        st.error("Acesso negado. Por favor, faça o login na Página Principal.")
        st.page_link("Home.py", label="Ir para a página de Login", icon="🏠")
        st.stop()
    
    if required_role and st.session_state.get('role') != required_role:
        st.error("Acesso negado. Você não tem permissão para ver esta página.")
        st.page_link("Home.py", label="Voltar à Página Principal", icon="🏠")
        st.stop()