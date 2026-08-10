
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Painel Comercial", page_icon="📊", layout="wide")

# ===== Identidade visual =====
NAVY = "#102A43"
BLUE = "#1F5A94"
LIGHT = "#EAF2F8"
GREEN = "#1E8449"
RED = "#C0392B"
GRAY = "#5D6D7E"

st.markdown(f"""
<style>
.stApp {{ background: #F6F8FB; }}
.block-container {{ padding-top: 1.2rem; max-width: 1400px; }}
h1,h2,h3 {{ color: {NAVY}; }}
.section {{
    background: white; padding: 18px 20px; border-radius: 12px;
    border: 1px solid #D9E2EC; margin: 12px 0;
}}
.header {{
    background: linear-gradient(90deg, {NAVY}, {BLUE});
    color: white; padding: 22px 28px; border-radius: 14px; margin-bottom: 18px;
}}
.metric-card {{
    background:white; border:1px solid #D9E2EC; border-radius:12px;
    padding:15px; text-align:center;
}}
.small {{ color:{GRAY}; font-size:0.9rem; }}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header">
<h1 style="color:white;margin:0;">📊 APRESENTAÇÃO — TIME COMERCIAL</h1>
<p style="margin:6px 0 0;">Painel de preenchimento e acompanhamento comercial</p>
</div>
""", unsafe_allow_html=True)

# ===== Estado =====
if "clientes" not in st.session_state:
    st.session_state.clientes = pd.DataFrame([
        {"CLIENTE":"","RECEITA":0.0,"MÊS ANTERIOR":0.0,"ANO ANTERIOR":0.0,
         "RENTABILIDADE":0.0,"SLA":0.0,"CONSIDERAÇÕES":""}
        for _ in range(5)
    ])
if "fechados" not in st.session_state:
    st.session_state.fechados = pd.DataFrame([
        {"CLIENTE":"","PROJEÇÃO MÊS":0.0,"FATURADO MÊS":0.0,"PRODUTO":"","ORIGEM":"","DESTINO":""}
        for _ in range(5)
    ])
if "upcoming" not in st.session_state:
    st.session_state.upcoming = pd.DataFrame([
        {"CLIENTE":"","PROJEÇÃO MÊS":0.0,"DATA PREVISTA":"","PRODUTO":"","ORIGEM":"","DESTINO":""}
        for _ in range(5)
    ])

# ===== Identificação =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("1. Identificação")
c1, c2 = st.columns(2)
with c1:
    nome = st.text_input("NOME")
with c2:
    mes = st.selectbox("MÊS", [
        "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
        "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
    ])
st.markdown('</div>', unsafe_allow_html=True)

# ===== Meta =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("2. META | FATURAMENTO")
c1,c2,c3,c4 = st.columns(4)
with c1: meta = st.number_input("META (R$)", min_value=0.0, step=1000.0)
with c2: faturado = st.number_input("FATURADO (R$)", min_value=0.0, step=1000.0)
with c3: projecao = st.number_input("PROJEÇÃO (R$)", min_value=0.0, step=1000.0)
with c4:
    ating = (faturado/meta*100) if meta else 0
    st.metric("% REALIZADO", f"{ating:.1f}%")
c1,c2 = st.columns(2)
with c1: st.metric("PROJEÇÃO %", f"{(projecao/meta*100 if meta else 0):.1f}%")
with c2: st.metric("GAP META x PROJEÇÃO", f"R$ {projecao-meta:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
st.markdown('</div>', unsafe_allow_html=True)

# ===== Histórico =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("3. META | FATURAMENTO — EVOLUÇÃO")
months = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
hist = pd.DataFrame({"Mês":months,
                     "Meta Total":[0.0]*12,
                     "Fat. Total":[0.0]*12,
                     "% Total":[0.0]*12,
                     "UP":[0.0]*12,
                     "LOSS":[0.0]*12})
edited_hist = st.data_editor(
    hist, use_container_width=True, hide_index=True,
    column_config={
        "Meta Total": st.column_config.NumberColumn(format="R$ %.2f"),
        "Fat. Total": st.column_config.NumberColumn(format="R$ %.2f"),
        "% Total": st.column_config.NumberColumn(format="%.1f%%"),
        "UP": st.column_config.NumberColumn(format="R$ %.2f"),
        "LOSS": st.column_config.NumberColumn(format="R$ %.2f"),
    },
    disabled=["Mês","% Total"], key="historico"
)
calc = edited_hist.copy()
calc["% Total"] = np.where(calc["Meta Total"]>0, calc["Fat. Total"]/calc["Meta Total"], 0)
calc["UP"] = np.maximum(calc["Fat. Total"]-calc["Meta Total"],0)
calc["LOSS"] = np.maximum(calc["Meta Total"]-calc["Fat. Total"],0)
st.line_chart(calc.set_index("Mês")[["Meta Total","Fat. Total"]])
st.markdown('</div>', unsafe_allow_html=True)

# ===== Clientes =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("4. PRINCIPAIS CLIENTES")
st.session_state.clientes = st.data_editor(
    st.session_state.clientes, num_rows="fixed", use_container_width=True,
    hide_index=True, key="clientes_editor",
    column_config={
        "RECEITA": st.column_config.NumberColumn(format="R$ %.2f"),
        "MÊS ANTERIOR": st.column_config.NumberColumn(format="R$ %.2f"),
        "ANO ANTERIOR": st.column_config.NumberColumn(format="R$ %.2f"),
        "RENTABILIDADE": st.column_config.NumberColumn(format="%.1f%%"),
        "SLA": st.column_config.NumberColumn(format="%.1f%%"),
    }
)
valid = st.session_state.clientes[st.session_state.clientes["CLIENTE"].astype(str).str.strip()!=""]
st.metric("TOTAL CARTEIRA", f"R$ {valid['RECEITA'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
st.markdown('</div>', unsafe_allow_html=True)

# ===== Pipeline =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("5. PIPELINE — NOVOS NEGÓCIOS FECHADOS - MÊS")
st.session_state.fechados = st.data_editor(
    st.session_state.fechados, num_rows="fixed", use_container_width=True,
    hide_index=True, key="fechados_editor",
    column_config={
        "PROJEÇÃO MÊS": st.column_config.NumberColumn(format="R$ %.2f"),
        "FATURADO MÊS": st.column_config.NumberColumn(format="R$ %.2f"),
    }
)
st.markdown("### UPCOMING – FECHADO PARA INICIAR")
st.session_state.upcoming = st.data_editor(
    st.session_state.upcoming, num_rows="fixed", use_container_width=True,
    hide_index=True, key="upcoming_editor",
    column_config={"PROJEÇÃO MÊS": st.column_config.NumberColumn(format="R$ %.2f")}
)
st.markdown('</div>', unsafe_allow_html=True)

# ===== Exportação Excel =====
st.markdown('<div class="section">', unsafe_allow_html=True)
st.subheader("6. EXPORTAR DADOS")
st.caption("O botão abaixo gera um Excel com as informações preenchidas. A geração do PowerPoint pode ser adicionada na próxima etapa.")
if st.button("📥 Gerar Excel"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([{
            "Nome":nome, "Mês":mes, "Meta":meta,
            "Faturado":faturado, "Projeção":projecao
        }]).to_excel(writer, sheet_name="Resumo", index=False)
        calc.to_excel(writer, sheet_name="Historico", index=False)
        st.session_state.clientes.to_excel(writer, sheet_name="Principais Clientes", index=False)
        st.session_state.fechados.to_excel(writer, sheet_name="Negocios Fechados", index=False)
        st.session_state.upcoming.to_excel(writer, sheet_name="Upcoming", index=False)
    st.download_button("⬇️ Baixar Excel", output.getvalue(),
                       file_name=f"painel_comercial_{mes.lower()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.markdown('</div>', unsafe_allow_html=True)

st.caption("Modelo inicial baseado na estrutura da apresentação comercial enviada.")
