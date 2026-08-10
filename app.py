import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.graph_objects as go

# ===== CONFIGURAÇÃO DA PÁGINA =====
st.set_page_config(page_title="Apresentação Comercial - KAO", page_icon="🚚", layout="wide")

# ===== ESTILIZACÃO CSS PARA PARECER SLIDES =====
st.markdown("""
<style>
/* Estilo das abas estilo PowerPoint */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #1E293B;
    padding: 10px;
    border-radius: 10px;
}

.stTabs [data-baseweb="tab"] {
    height: 45px;
    white-space: pre-wrap;
    background-color: #0F172A;
    border-radius: 8px;
    color: #94A3B8;
    font-weight: 600;
}

.stTabs [aria-selected="true"] {
    background-color: #22C55E !important;
    color: #000000 !important;
    font-weight: 800;
}
</style>
""", unsafe_allow_html=True)

# ===== CRIAÇÃO DAS PÁGINAS / SLIDES =====
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "1. Identificação",
    "2. Meta & Faturamento",
    "3. Evolução Histórica",
    "4. Top 5 Clientes",
    "5. Novos Fechados",
    "6. Pipeline",
    "7. Exportação"
])

# ----------------------------------------------------
# SLIDE 1: IDENTIFICAÇÃO
# ----------------------------------------------------
with tab1:
    st.subheader("1. Identificação do Executivo")
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        nome = st.text_input("NOME DO EXECUTIVO / KAM", value="Cleyton Kauffmann")
    with c2:
        mes = st.selectbox("MÊS DE REFERÊNCIA", ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], index=7)
    with c3:
        data_ref = st.date_input("DATA DE APRESENTAÇÃO")

# ----------------------------------------------------
# SLIDE 2: META | FATURAMENTO
# ----------------------------------------------------
with tab2:
    st.subheader("2. META | FATURAMENTO — MÊS ATUAL")
    col_meta, col_fat, col_proj = st.columns(3)
    with col_meta:
        meta = st.number_input("📌 META DA GERÊNCIA (R$)", min_value=0.0, value=382658.00, step=1000.0)
    with col_fat:
        faturado = st.number_input("💰 FATURADO ALCANÇADO (R$)", min_value=0.0, value=382000.00, step=1000.0)
    with col_proj:
        projecao = st.number_input("📈 PROJEÇÃO MÊS (R$)", min_value=0.0, value=faturado, step=1000.0)
        
    # [Seus cards com os indicadores circulares entram aqui]

# ----------------------------------------------------
# SLIDE 3: EVOLUÇÃO HISTÓRICA
# ----------------------------------------------------
with tab3:
    st.subheader("3. META | FATURAMENTO — EVOLUÇÃO HISTÓRICA")
    # [Coloque aqui a tabela st.data_editor do histórico e o gráfico do Plotly]

# ----------------------------------------------------
# SLIDE 4: TOP 5 CLIENTES
# ----------------------------------------------------
with tab4:
    st.subheader("4. PRINCIPAIS CLIENTES (TOP 5)")
    # [Coloque aqui a tabela dos clientes]

# ----------------------------------------------------
# SLIDE 5: NOVOS FECHADOS
# ----------------------------------------------------
with tab5:
    st.subheader("5. NOVOS CLIENTES FECHADOS NO MÊS")
    # [Coloque aqui a tabela de novos fechados]

# ----------------------------------------------------
# SLIDE 6: PIPELINE
# ----------------------------------------------------
with tab6:
    st.subheader("6. PRÓXIMOS FECHAMENTOS (PIPELINE / UPCOMING)")
    # [Coloque aqui a tabela de próximos fechamentos]

# ----------------------------------------------------
# SLIDE 7: EXPORTAÇÃO
# ----------------------------------------------------
with tab7:
    st.subheader("7. Exportação dos Dados")
    # [Coloque aqui o botão de download do Excel]
