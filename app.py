import streamlit as st
import os

# 1. Configuração da página (deve ser o primeiro comando)
st.set_page_config(
    page_title="Briefings ComexStat",
    page_icon="📊",
    layout="wide"
)

# 2. Logo da Sidebar (colocada aqui, ela fica ACIMA da navegação)
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

# 3. O Streamlit cuida do resto
# Ele irá automaticamente encontrar a pasta 'pages/' e criar a navegação
# abaixo da logo. A página '0_Home.py' será listada como 'Home'.

# --- O conteúdo da página Home foi movido para 'pages/0_Home.py' ---
# --- Este arquivo agora controla apenas a configuração global ---
