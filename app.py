import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.graph_objects as go

# ===== FUNÇÃO DE FORMATAÇÃO BRASILEIRA =====
def fmt_br(valor):
    try:
        return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "0,00"

# ===== CONFIGURAÇÃO DA PÁGINA =====
st.set_page_config(page_title="Apresentação Comercial - KAO", page_icon="🚚", layout="wide")

# ===== IDENTIDADE VISUAL =====
BG_DARK = "#0F172A"       
CARD_DARK = "#1E293B"     
BORDER_DARK = "#334155"   
TEXT_MAIN = "#F8FAFC"     
TEXT_MUTED = "#94A3B8"    
GREEN_NEON = "#22C55E"    
NAVY_ACCENT = "#38BDF8"   

st.markdown(f"""
<style>
.stApp {{ background-color: {BG_DARK}; color: {TEXT_MAIN}; }}
.block-container {{ padding-top: 1.5rem; max-width: 1350px; }}
h1, h2, h3, h4, h5 {{ color: {TEXT_MAIN} !important; font-family: 'Segoe UI', Roboto, sans-serif; }}

.header-slide {{
    background: linear-gradient(135deg, #0284C7 0%, #0F172A 100%);
    border-left: 8px solid {GREEN_NEON};
    color: white;
    padding: 22px 28px;
    border-radius: 12px;
    margin-bottom: 24px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.5);
}}

.section-card {{
    background: {CARD_DARK};
    padding: 24px;
    border-radius: 14px;
    border: 1px solid {BORDER_DARK};
    margin-bottom: 24px;
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.4);
}}

.stTextInput label, .stSelectbox label, .stNumberInput label, .stDateInput label {{
    color: {TEXT_MAIN} !important;
    font-weight: 600;
}}

[data-testid="stDataEditor"] {{
    background-color: {CARD_DARK};
    border-radius: 8px;
}}
</style>
""", unsafe_allow_html=True)

# ===== HEADER PRINCIPAL =====
st.markdown(f"""
<div class="header-slide">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h1 style="color:white !important; margin:0; font-size:1.9rem; font-weight:700;">📊 APRESENTAÇÃO — TIME COMERCIAL</h1>
            <p style="margin:6px 0 0; color:#E2E8F0; font-size:1.0rem;">Painel de Acompanhamento e Desempenho Operacional / Comercial</p>
        </div>
        <div style="text-align:right;">
            <span style="background:{GREEN_NEON}; color:#000; padding:8px 16px; border-radius:20px; font-weight:800; font-size:0.85rem; letter-spacing:0.5px;">SISTEMA KAO</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ===== ESTADO DA SESSÃO (INICIALIZAÇÃO) =====
months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
          "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

if "df_historico" not in st.session_state:
    st.session_state.df_historico = pd.DataFrame({
        "Mês": months,
        "Meta Total": [382658.00] * 12,
        "Fat. Total": [0.0] * 12,
        "% Total": [0.0] * 12,
        "UP": [0.0] * 12,
        "LOSS": [0.0] * 12
    })

if "clientes" not in st.session_state:
    st.session_state.clientes = pd.DataFrame([
        {"CLIENTE": f"Cliente {i+1}", "RECEITA": 0.0, "MÊS ANTERIOR": 0.0, "ANO ANTERIOR": 0.0, "RENTABILIDADE": 0.0, "SLA": 0.0, "CONSIDERAÇÕES": ""}
        for i in range(5)
    ])

if "fechados" not in st.session_state:
    st.session_state.fechados = pd.DataFrame([
        {"CLIENTE": "", "PROJEÇÃO MÊS": 0.0, "FATURADO MÊS": 0.0, "PRODUTO": "", "ORIGEM": "", "DESTINO": ""}
        for _ in range(5)
    ])

if "upcoming" not in st.session_state:
    st.session_state.upcoming = pd.DataFrame([
        {"CLIENTE": "", "PROJEÇÃO MÊS": 0.0, "DATA PREVISTA": "", "PRODUTO": "", "ORIGEM": "", "DESTINO": ""}
        for _ in range(5)
    ])

if "faturado_auto" not in st.session_state:
    st.session_state.faturado_auto = 88873.92

# ===== IMPORTAÇÃO DA PLANILHA =====
with st.expander("📂 Importar Dados da Planilha (.xlsx)", expanded=True):
    uploaded_file = st.file_uploader("Arraste ou selecione a planilha (ex: data (50).xlsx)", type=["xlsx"])
    if uploaded_file:
        try:
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            sheet_target = "Export" if "Export" in excel_data else list(excel_data.keys())[0]
            df_export = excel_data[sheet_target]
            
            if "Razão Grupo" in df_export.columns and "Mês atual" in df_export.columns:
                df_valid = df_export[
                    df_export["Razão Grupo"].notna() & 
                    (~df_export["Razão Grupo"].astype(str).str.startswith("Total")) &
                    (~df_export["Razão Grupo"].astype(str).str.startswith("Filtros"))
                ].copy()

                df_valid["Mês atual"] = pd.to_numeric(df_valid["Mês atual"], errors='coerce').fillna(0.0)
                df_valid["Mês anterior"] = pd.to_numeric(df_valid["Mês anterior"], errors='coerce').fillna(0.0)

                total_faturado_calc = float(df_valid["Mês atual"].sum())
                
                if st.session_state.faturado_auto != total_faturado_calc:
                    st.session_state.faturado_auto = total_faturado_calc

                    top5 = df_valid.sort_values(by="Mês atual", ascending=False).head(5)
                    novos_clientes = []
                    for idx, row in top5.iterrows():
                        novos_clientes.append({
                            "CLIENTE": str(row["Razão Grupo"]),
                            "RECEITA": float(row["Mês atual"]),
                            "MÊS ANTERIOR": float(row["Mês anterior"]),
                            "ANO ANTERIOR": 0.0,
                            "RENTABILIDADE": 0.0,
                            "SLA": 0.0,
                            "CONSIDERAÇÕES": ""
                        })

                    while len(novos_clientes) < 5:
                        novos_clientes.append({
                            "CLIENTE": "", "RECEITA": 0.0, "MÊS ANTERIOR": 0.0,
                            "ANO ANTERIOR": 0.0, "RENTABILIDADE": 0.0, "SLA": 0.0, "CONSIDERAÇÕES": ""
                        })

                    st.session_state.clientes = pd.DataFrame(novos_clientes)
                    st.success(f"✅ Sucesso! Faturado total de R$ {fmt_br(total_faturado_calc)} importado!")
                    st.rerun()
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

# ===== SEÇÃO 1: IDENTIFICAÇÃO =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("1. Identificação do Executivo")
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    nome = st.text_input("NOME DO EXECUTIVO / KAM", value="Cleyton Kauffmann")
with c2:
    mes = st.selectbox("MÊS DE REFERÊNCIA", months, index=7)
with c3:
    data_ref = st.date_input("DATA DE APRESENTAÇÃO")
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 2: META | FATURAMENTO =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("2. META | FATURAMENTO — MÊS ATUAL")

col_meta, col_fat, col_proj = st.columns(3)
with col_meta:
    meta = st.number_input("📌 META DA GERÊNCIA (R$)", min_value=0.0, value=382658.00, step=1000.0)
with col_fat:
    faturado = st.number_input("💰 FATURADO ALCANÇADO (R$)", min_value=0.0, value=float(st.session_state.faturado_auto), step=1000.0)
with col_proj:
    projecao = st.number_input("📈 PROJEÇÃO MÊS (R$)", min_value=0.0, value=faturado, step=1000.0)

pct_realizado = (faturado / meta * 100) if meta > 0 else 0.0
pct_projecao = (projecao / meta * 100) if meta > 0 else 0.0
gap_projecao = projecao - meta
delta_pct = pct_realizado - 100

st.divider()

m1, m2, m3 = st.columns(3)
cor_bola_bg = "linear-gradient(135deg, #166534 0%, #22C55E 100%)" if pct_realizado >= 100 else "linear-gradient(135deg, #991B1B 0%, #EF4444 100%)"
cor_badge_bg = "#14532D" if delta_pct >= 0 else "#7F1D1D"
cor_badge_txt = "#4ADE80" if delta_pct >= 0 else "#FCA5A5"

with m1:
    st.markdown(f"""
    <div style="text-align: center; background: #0F172A; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
        <span style="font-size: 0.85rem; font-weight: 700; color: {TEXT_MUTED}; text-transform: uppercase;">% ALCANÇADO (META x FAT)</span>
        <div style="width: 120px; height: 120px; background: {cor_bola_bg}; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 15px auto;">
            <span style="color: #FFFFFF; font-size: 1.6rem; font-weight: 800;">{pct_realizado:.1f}%</span>
        </div>
        <span style="background-color: {cor_badge_bg}; color: {cor_badge_txt}; padding: 6px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700;">
            {"↑" if delta_pct >= 0 else "↓"} {delta_pct:+.1f}% vs Meta
        </span>
    </div>
    """.replace(".", ","), unsafe_allow_html=True)

with m2:
    st.markdown(f"""
    <div style="text-align: center; background: #0F172A; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
        <span style="font-size: 0.85rem; font-weight: 700; color: {TEXT_MUTED}; text-transform: uppercase;">PROJEÇÃO DA META %</span>
        <div style="width: 120px; height: 120px; background: linear-gradient(135deg, #0369A1 0%, #38BDF8 100%); border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 15px auto;">
            <span style="color: #FFFFFF; font-size: 1.6rem; font-weight: 800;">{pct_projecao:.1f}%</span>
        </div>
        <span style="color: {TEXT_MUTED}; font-size: 0.85rem; font-weight: 600;">Projeção de Fechamento</span>
    </div>
    """.replace(".", ","), unsafe_allow_html=True)

with m3:
    cor_gap = "#4ADE80" if gap_projecao >= 0 else "#FCA5A5"
    st.markdown(f"""
    <div style="text-align: center; background: #0F172A; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
        <span style="font-size: 0.85rem; font-weight: 700; color: {TEXT_MUTED}; text-transform: uppercase;">GAP (PROJEÇÃO x META)</span>
        <div style="height: 120px; display: flex; align-items: center; justify-content: center; margin: 15px auto;">
            <span style="color: #FFFFFF; font-size: 2.0rem; font-weight: 800;">R$ {fmt_br(gap_projecao)}</span>
        </div>
        <span style="background-color: {cor_badge_bg}; color: {cor_gap}; padding: 6px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700;">
            {"↑" if gap_projecao >= 0 else "↓"} R$ {fmt_br(gap_projecao)}
        </span>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 3: EVOLUÇÃO HISTÓRICA =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("3. META | FATURAMENTO — EVOLUÇÃO HISTÓRICA")

idx_mes = months.index(mes)
st.session_state.df_historico.loc[idx_mes, "Fat. Total"] = faturado

# Recalcular métricas
df_calc = st.session_state.df_historico.copy()
df_calc["% Total"] = np.where(df_calc["Meta Total"] > 0, (df_calc["Fat. Total"] / df_calc["Meta Total"]) * 100, 0.0)
df_calc["UP"] = np.maximum(df_calc["Fat. Total"] - df_calc["Meta Total"], 0)
df_calc["LOSS"] = np.maximum(df_calc["Meta Total"] - df_calc["Fat. Total"], 0)
st.session_state.df_historico = df_calc

# Função para formatação condicional de cor
def colorir_atingido(val):
    if val >= 100.0:
        return "color: #22C55E; font-weight: bold; background-color: rgba(34, 197, 94, 0.15);" # Verde
    elif val > 0:
        return "color: #EF4444; font-weight: bold; background-color: rgba(239, 68, 68, 0.15);"  # Vermelho
    return "color: #94A3B8;" # Cinza quando zerado

# Aplicação do Styler na tabela
try:
    df_estilizado = st.session_state.df_historico.style.map(
        colorir_atingido, subset=["% Total"]
    ).format({
        "Meta Total": "R$ {:,.2f}",
        "Fat. Total": "R$ {:,.2f}",
        "% Total": "{:.1f}%",
        "UP": "R$ {:,.2f}",
        "LOSS": "R$ {:,.2f}"
    })
except AttributeError:
    df_estilizado = st.session_state.df_historico.style.applymap(
        colorir_atingido, subset=["% Total"]
    ).format({
        "Meta Total": "R$ {:,.2f}",
        "Fat. Total": "R$ {:,.2f}",
        "% Total": "{:.1f}%",
        "UP": "R$ {:,.2f}",
        "LOSS": "R$ {:,.2f}"
    })

st.dataframe(
    df_estilizado,
    use_container_width=True,
    hide_index=True
)

# Gráfico da evolução
fig = go.Figure()
fig.add_trace(go.Bar(
    x=st.session_state.df_historico["Mês"],
    y=st.session_state.df_historico["Fat. Total"],
    name="Faturado Alcançado",
    marker_color=GREEN_NEON,
    width=0.35
))
fig.add_trace(go.Scatter(
    x=st.session_state.df_historico["Mês"],
    y=st.session_state.df_historico["Fat. Total"],
    name="Tendência",
    mode="lines+markers",
    line=dict(color=NAVY_ACCENT, dash="dash", width=3)
))
fig.update_layout(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0F172A",
    font=dict(color="#F8FAFC"), margin=dict(l=20, r=20, t=20, b=20)
)
st.plotly_chart(fig, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 4: TOP 5 CLIENTES =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("4. PRINCIPAIS CLIENTES (TOP 5)")

st.session_state.clientes = st.data_editor(
    st.session_state.clientes,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    key="editor_clientes_v2",
    column_config={
        "RECEITA": st.column_config.NumberColumn("RECEITA (R$)", format="R$ %,.2f"),
        "MÊS ANTERIOR": st.column_config.NumberColumn("MÊS ANTERIOR (R$)", format="R$ %,.2f"),
        "ANO ANTERIOR": st.column_config.NumberColumn("ANO ANTERIOR (R$)", format="R$ %,.2f"),
        "RENTABILIDADE": st.column_config.NumberColumn("RENTABILIDADE %", format="%.1f%%"),
        "SLA": st.column_config.NumberColumn("SLA %", format="%.1f%%"),
        "CONSIDERAÇÕES": st.column_config.TextColumn("CONSIDERAÇÕES / OBS")
    }
)
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 5: NOVOS CLIENTES FECHADOS =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("5. NOVOS CLIENTES FECHADOS NO MÊS")

st.session_state.fechados = st.data_editor(
    st.session_state.fechados,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="editor_fechados_v2",
    column_config={
        "PROJEÇÃO MÊS": st.column_config.NumberColumn("PROJEÇÃO (R$)", format="R$ %,.2f"),
        "FATURADO MÊS": st.column_config.NumberColumn("FATURADO (R$)", format="R$ %,.2f"),
    }
)
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 6: PRÓXIMOS FECHAMENTOS (PIPELINE) =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("6. PRÓXIMOS FECHAMENTOS (PIPELINE / UPCOMING)")

st.session_state.upcoming = st.data_editor(
    st.session_state.upcoming,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="editor_upcoming_v2",
    column_config={
        "PROJEÇÃO MÊS": st.column_config.NumberColumn("PROJEÇÃO (R$)", format="R$ %,.2f"),
    }
)
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 7: EXPORTAÇÃO =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("7. Exportação dos Dados")

buffer_excel = BytesIO()
with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
    st.session_state.df_historico.to_excel(writer, sheet_name="Historico", index=False)
    st.session_state.clientes.to_excel(writer, sheet_name="Top 5 Clientes", index=False)
    st.session_state.fechados.to_excel(writer, sheet_name="Fechados", index=False)
    st.session_state.upcoming.to_excel(writer, sheet_name="Pipeline", index=False)

st.download_button(
    label="📊 Baixar Relatório Consolidado em Excel",
    data=buffer_excel.getvalue(),
    file_name=f"Relatorio_Comercial_{mes}_{nome.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.markdown('</div>', unsafe_allow_html=True)
