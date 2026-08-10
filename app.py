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

# ===== IDENTIDADE VISUAL & CSS ESTILO SLIDES =====
BG_DARK = "#0F172A"       
CARD_DARK = "#1E293B"     
BORDER_DARK = "#334155"   
TEXT_MAIN = "#F8FAFC"     
TEXT_MUTED = "#94A3B8"    
GREEN_NEON = "#22C55E"    
NAVY_ACCENT = "#38BDF8"   

st.markdown(f"""
<style>
/* Definição visual para títulos estilizados */
.stApp {{ background-color: {BG_DARK}; color: {TEXT_MAIN}; }}
.block-container {{ padding-top: 1.5rem; max-width: 1350px; }}
h1, h2, h3, h4, h5 {{ color: {TEXT_MAIN} !important; font-family: 'Amasis MT Pro', 'Georgia', serif; }}

/* Fonte Personalizada Executiva Imponente */
.executive-font {{
    font-family: 'Amasis MT Pro Black', 'Amasis MT Pro', 'Georgia', serif !important;
    font-weight: 900 !important;
    letter-spacing: 0.5px;
}}

/* Header Principal */
.header-slide {{
    background: linear-gradient(135deg, #0284C7 0%, #0F172A 100%);
    border-left: 8px solid {GREEN_NEON};
    color: white;
    padding: 22px 28px;
    border-radius: 12px;
    margin-bottom: 24px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.5);
}}

/* Estilo das Abas de Navegação (PowerPoint Style) */
.stTabs [data-baseweb="tab-list"] {{
    gap: 8px;
    background-color: {CARD_DARK};
    padding: 10px;
    border-radius: 10px;
    border: 1px solid {BORDER_DARK};
}}

.stTabs [data-baseweb="tab"] {{
    height: 45px;
    white-space: pre-wrap;
    background-color: {BG_DARK};
    border-radius: 8px;
    color: {TEXT_MUTED};
    font-weight: 600;
}}

.stTabs [aria-selected="true"] {{
    background-color: {GREEN_NEON} !important;
    color: #000000 !important;
    font-weight: 800;
}}

/* Foto de Perfil Arredondada */
.profile-img {{
    width: 180px;
    height: 180px;
    border-radius: 50%;
    object-fit: cover;
    border: 4px solid {GREEN_NEON};
    box-shadow: 0px 8px 20px rgba(0,0,0,0.4);
}}

/* Card da Identificação Clean */
.id-card {{
    background-color: {CARD_DARK};
    border: 1px solid {BORDER_DARK};
    padding: 30px 35px;
    border-radius: 14px;
    box-shadow: 0 8px 25px rgba(0,0,0,0.3);
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
    st.session_state.faturado_auto = 382000.00

# ===== BARRA LATERAL (SIDEBAR): TODOS OS FORMULÁRIOS E UPLOADS =====
with st.sidebar:
    st.header("⚙️ Painel de Controle")
    
    with st.expander("👤 Dados da Identificação", expanded=True):
        nome_input = st.text_input("Nome do Executivo / KAM", value="Cleyton Kauffmann")
        mes_input = st.selectbox("Mês de Referência", months, index=7)
        data_ref = st.date_input("Data de Apresentação")
        foto_upload = st.file_uploader("Foto de Perfil", type=["png", "jpg", "jpeg"])

    with st.expander("📂 Importar Planilha Excel", expanded=False):
        uploaded_file = st.file_uploader("Arquivo .xlsx", type=["xlsx"])

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
                        st.success(f"✅ Dados atualizados!")
                        st.rerun()
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

# ===== HEADER PRINCIPAL =====
st.markdown(f"""
<div class="header-slide">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h1 class="executive-font" style="color:white !important; margin:0; font-size:2.0rem;">APRESENTAÇÃO — TIME COMERCIAL</h1>
            <p style="margin:6px 0 0; color:#E2E8F0; font-size:1.0rem;">Painel de Acompanhamento e Desempenho Operacional / Comercial</p>
        </div>
        <div style="text-align:right;">
            <span style="background:{GREEN_NEON}; color:#000; padding:8px 16px; border-radius:20px; font-weight:800; font-size:0.85rem; letter-spacing:0.5px;">SISTEMA KAO</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ===== CRIAÇÃO DAS ABAS (SLIDES) =====
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
# SLIDE 1: IDENTIFICAÇÃO (APENAS O CARD VISUAL EXECUTIVO)
# ----------------------------------------------------
with tab1:
    st.subheader("1. Identificação do Executivo")
    st.write("")
    
    col_foto, col_info = st.columns([1, 3])
    
    with col_foto:
        if foto_upload is not None:
            st.image(foto_upload, width=180)
        else:
            st.markdown(f"""
            <div style="
                width: 180px; height: 180px; border-radius: 50%; 
                background-color: {CARD_DARK}; border: 3px dashed {BORDER_DARK}; 
                display: flex; align-items: center; justify-content: center; 
                color: {TEXT_MUTED}; font-size: 0.85rem; text-align: center;">
                📸 Envie a foto na barra lateral
            </div>
            """, unsafe_allow_html=True)

    with col_info:
        # Card Visual com Fonte Amasis MT Pro Black
        st.markdown(f"""
        <div class="id-card">
            <span style="color: {TEXT_MUTED}; font-size: 0.9rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;">
                Executivo de Contas / KAM
            </span>
            <div class="executive-font" style="font-size: 2.6rem; color: #FFFFFF; margin: 6px 0 12px 0; text-transform: uppercase;">
                {nome_input}
            </div>
            <div style="display: flex; gap: 15px; align-items: center;">
                <span class="executive-font" style="background-color: {GREEN_NEON}; color: #000; padding: 6px 16px; border-radius: 8px; font-size: 1.1rem;">
                    MÊS: {mes_input.upper()}
                </span>
                <span style="color: {NAVY_ACCENT}; font-weight: 600; font-size: 1.05rem;">
                    📅 Apresentação: {data_ref.strftime('%d/%m/%Y')}
                </span>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ----------------------------------------------------
# SLIDE 2: META & FATURAMENTO
# ----------------------------------------------------
with tab2:
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
        <div style="text-align: center; background: #1E293B; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
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
        <div style="text-align: center; background: #1E293B; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
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
        <div style="text-align: center; background: #1E293B; padding: 20px; border-radius: 14px; border: 1px solid {BORDER_DARK};">
            <span style="font-size: 0.85rem; font-weight: 700; color: {TEXT_MUTED}; text-transform: uppercase;">GAP (PROJEÇÃO x META)</span>
            <div style="height: 120px; display: flex; align-items: center; justify-content: center; margin: 15px auto;">
                <span style="color: #FFFFFF; font-size: 2.0rem; font-weight: 800;">R$ {fmt_br(gap_projecao)}</span>
            </div>
            <span style="background-color: {cor_badge_bg}; color: {cor_gap}; padding: 6px 14px; border-radius: 20px; font-size: 0.85rem; font-weight: 700;">
                {"↑" if gap_projecao >= 0 else "↓"} R$ {fmt_br(gap_projecao)}
            </span>
        </div>
        """, unsafe_allow_html=True)

# ----------------------------------------------------
# SLIDE 3: EVOLUÇÃO HISTÓRICA
# ----------------------------------------------------
with tab3:
    st.subheader("3. META | FATURAMENTO — EVOLUÇÃO HISTÓRICA")

    idx_mes = months.index(mes_input)
    st.session_state.df_historico.loc[idx_mes, "Fat. Total"] = faturado

    # Recalcular métricas
    df_calc = st.session_state.df_historico.copy()
    df_calc["% Total"] = np.where(df_calc["Meta Total"] > 0, (df_calc["Fat. Total"] / df_calc["Meta Total"]) * 100, 0.0)
    df_calc["UP"] = np.maximum(df_calc["Fat. Total"] - df_calc["Meta Total"], 0)
    df_calc["LOSS"] = np.maximum(df_calc["Meta Total"] - df_calc["Fat. Total"], 0)
    st.session_state.df_historico = df_calc

    # Tabela Editável
    st.session_state.df_historico = st.data_editor(
        st.session_state.df_historico,
        use_container_width=True,
        hide_index=True,
        key="editor_historico_tabs",
        column_config={
            "Mês": st.column_config.TextColumn("Mês", disabled=True),
            "Meta Total": st.column_config.NumberColumn("Meta Total (R$)", format="R$ %,.2f"),
            "Fat. Total": st.column_config.NumberColumn("Fat. Total (R$)", format="R$ %,.2f"),
            "% Total": st.column_config.NumberColumn("% Atingido", format="%.1f%%", disabled=True),
            "UP": st.column_config.NumberColumn("UP (R$)", format="R$ %,.2f", disabled=True),
            "LOSS": st.column_config.NumberColumn("LOSS (R$)", format="R$ %,.2f", disabled=True),
        }
    )

    # Gráfico Ajustado (Com rótulos na base)
    valores_formatados = [f"R$ {fmt_br(v)}" if v > 0 else "" for v in st.session_state.df_historico["Fat. Total"]]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=st.session_state.df_historico["Mês"],
        y=st.session_state.df_historico["Fat. Total"],
        name="Faturado Alcançado",
        marker_color=GREEN_NEON,
        text=valores_formatados,
        textposition="inside",
        insidetextanchor="end",
        width=0.4,
        hovertemplate="<b>%{x}</b><br>Faturado: R$ %{y:,.2f}<extra></extra>"
    ))

    fig.add_trace(go.Scatter(
        x=st.session_state.df_historico["Mês"],
        y=st.session_state.df_historico["Fat. Total"],
        name="Tendência",
        mode="lines+markers",
        line=dict(color=NAVY_ACCENT, dash="dash", width=3),
        hovertemplate="<b>%{x}</b><br>Tendência: R$ %{y:,.2f}<extra></extra>"
    ))

    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0F172A",
        font=dict(color="#F8FAFC"), margin=dict(l=20, r=20, t=30, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5)
    )

    st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------
# SLIDE 4: TOP 5 CLIENTES
# ----------------------------------------------------
with tab4:
    st.subheader("4. PRINCIPAIS CLIENTES (TOP 5)")

    st.session_state.clientes = st.data_editor(
        st.session_state.clientes,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key="editor_clientes_tabs",
        column_config={
            "RECEITA": st.column_config.NumberColumn("RECEITA (R$)", format="R$ %,.2f"),
            "MÊS ANTERIOR": st.column_config.NumberColumn("MÊS ANTERIOR (R$)", format="R$ %,.2f"),
            "ANO ANTERIOR": st.column_config.NumberColumn("ANO ANTERIOR (R$)", format="R$ %,.2f"),
            "RENTABILIDADE": st.column_config.NumberColumn("RENTABILIDADE %", format="%.1f%%"),
            "SLA": st.column_config.NumberColumn("SLA %", format="%.1f%%"),
            "CONSIDERAÇÕES": st.column_config.TextColumn("CONSIDERAÇÕES / OBS")
        }
    )

# ----------------------------------------------------
# SLIDE 5: NOVOS CLIENTES FECHADOS
# ----------------------------------------------------
with tab5:
    st.subheader("5. NOVOS CLIENTES FECHADOS NO MÊS")

    st.session_state.fechados = st.data_editor(
        st.session_state.fechados,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_fechados_tabs",
        column_config={
            "PROJEÇÃO MÊS": st.column_config.NumberColumn("PROJEÇÃO (R$)", format="R$ %,.2f"),
            "FATURADO MÊS": st.column_config.NumberColumn("FATURADO (R$)", format="R$ %,.2f"),
        }
    )

# ----------------------------------------------------
# SLIDE 6: PIPELINE
# ----------------------------------------------------
with tab6:
    st.subheader("6. PRÓXIMOS FECHAMENTOS (PIPELINE / UPCOMING)")

    st.session_state.upcoming = st.data_editor(
        st.session_state.upcoming,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_upcoming_tabs",
        column_config={
            "PROJEÇÃO MÊS": st.column_config.NumberColumn("PROJEÇÃO (R$)", format="R$ %,.2f"),
        }
    )

# ----------------------------------------------------
# SLIDE 7: EXPORTAÇÃO
# ----------------------------------------------------
with tab7:
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
        file_name=f"Relatorio_Comercial_{mes_input}_{nome_input.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
