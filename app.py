import streamlit as st
import os

# 1. Configuração da página (deve ser o primeiro comando)
st.set_page_config(
    page_title="Central de Automações | AEST",
    page_icon="📊",
    layout="wide"
)

# 2. Logo da Sidebar (colocada aqui, ela fica ACIMA da navegação)
logo_sidebar_path = "LogoMinasGerais.png"
if os.path.exists(logo_sidebar_path):
    st.sidebar.image(logo_sidebar_path, width=200)

# 3. O Streamlit cuida do resto
# Ele irá automaticamente encontrar a pasta 'pages/' e criar a navegação
# abaixo da logo, começando com '0_Home.py'
