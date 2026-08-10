import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ===== Função de Formatação Brasileira (ex: 350.658,00) =====
def fmt_br(valor):
    try:
        return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "0,00"

# ===== Configuração da Página =====
st.set_page_config(page_title="Apresentação Comercial - KAO", page_icon="🚚", layout="wide")

# ===== Identidade visual (Alinhada aos slides) =====
NAVY = "#102A43"
GREEN = "#1E8449"
BLUE = "#1F5A94"
GRAY = "#5D6D7E"

st.markdown(f"""
<style>
.stApp {{ background-color: #F8FAFC; }}
.block-container {{ padding-top: 1.5rem; max-width: 1350px; }}
h1, h2, h3 {{ color: {NAVY}; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}

.header-slide {{
    background: linear-gradient(135deg, {NAVY} 0%, #163C66 100%);
    border-left: 8px solid {GREEN};
    color: white;
    padding: 20px 25px;
    border-radius: 10px;
    margin-bottom: 20px;
    box-shadow: 0 4px 6px rgba(0,0,0,0.05);
}}

.section-card {{
    background: white;
    padding: 20px;
    border-radius: 10px;
    border: 1px solid #E2E8F0;
    margin-bottom: 20px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.02);
}}
</style>
""", unsafe_allow_html=True)

# ===== Header Principal =====
st.markdown(f"""
<div class="header-slide">
    <div style="display:flex; justify-content:space-between; align-items:center;">
        <div>
            <h1 style="color:white; margin:0; font-size:1.8rem;">📊 APRESENTAÇÃO — TIME COMERCIAL</h1>
            <p style="margin:4px 0 0; opacity:0.85; font-size:0.95rem;">Painel de Acompanhamento e Desempenho Operacional / Comercial</p>
        </div>
        <div style="text-align:right;">
            <span style="background:{GREEN}; padding:6px 12px; border-radius:20px; font-weight:bold; font-size:0.85rem;">SISTEMA KAO</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ===== Estado da Sessão (Session State) =====
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

# ===== SEÇÃO 0: IMPORTAÇÃO DE DADOS (OPCIONAL) =====
with st.expander("📂 Importar Dados de Planilha Excel (Carga Rápida)", expanded=False):
    uploaded_file = st.file_uploader("Arraste ou selecione a planilha de metas/histórico (.xlsx)", type=["xlsx"])
    if uploaded_file:
        try:
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            st.success("Planilha carregada com sucesso!")
            if "Export" in excel_data:
                st.session_state["df_export_imported"] = excel_data["Export"]
            elif "Historico" in excel_data:
                st.session_state["df_hist_imported"] = excel_data["Historico"]
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

# ===== SEÇÃO 1: IDENTIFICAÇÃO =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("1. Identificação")
c1, c2, c3 = st.columns([2, 1, 1])
with c1:
    nome = st.text_input("NOME DO EXECUTIVO / KAM", value="Cleyton Kauffmann")
with c2:
    mes = st.selectbox("MÊS DE REFERÊNCIA", [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ], index=4)
with c3:
    data_ref = st.date_input("DATA DE APRESENTAÇÃO")
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 2: META | FATURAMENTO (MÊS ATUAL) =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("2. META | FATURAMENTO - MÊS ATUAL")
st.caption("A Meta é definida previamente pela Gerência. Insira o Faturado Alcançado para apurar o resultado.")

col_meta, col_fat, col_proj = st.columns(3)

with col_meta:
    meta = st.number_input("📌 META DA GERÊNCIA (R$)", min_value=0.0, value=350658.00, step=1000.0)

with col_fat:
    faturado = st.number_input("💰 FATURADO ALCANÇADO (R$)", min_value=0.0, value=0.0, step=1000.0)

with col_proj:
    projecao = st.number_input("📈 PROJEÇÃO MÊS (R$)", min_value=0.0, value=faturado, step=1000.0)

# CÁLCULOS
pct_realizado = (faturado / meta * 100) if meta > 0 else 0.0
pct_projecao = (projecao / meta * 100) if meta > 0 else 0.0
gap_projecao = projecao - meta

st.divider()

m1, m2, m3 = st.columns(3)
with m1:
    st.metric(
        label="% ALCANÇADO (META x FATURADO)",
        value=f"{pct_realizado:.1f}%".replace(".", ","),
        delta=f"{(pct_realizado - 100):.1f}% vs Meta".replace(".", ",")
    )
with m2:
    st.metric(
        label="PROJEÇÃO DA META %",
        value=f"{pct_projecao:.1f}%".replace(".", ",")
    )
with m3:
    st.metric(
        label="GAP (PROJEÇÃO x META)",
        value=f"R$ {fmt_br(gap_projecao)}",
        delta=f"R$ {fmt_br(gap_projecao)}"
    )
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 3: EVOLUÇÃO HISTÓRICA =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("3. META | FATURAMENTO — EVOLUÇÃO HISTÓRICA")

months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

if "df_hist_imported" in st.session_state:
    hist_base = st.session_state["df_hist_imported"]
else:
    hist_base = pd.DataFrame({
        "Mês": months,
        "Meta Total": [350658.00] * 12,
        "Fat. Total": [0.0] * 12,
        "% Total": [0.0] * 12,
        "UP": [0.0] * 12,
        "LOSS": [0.0] * 12
    })

edited_hist = st.data_editor(
    hist_base,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Mês": st.column_config.TextColumn(disabled=True),
        "Meta Total": st.column_config.NumberColumn("Meta Total (R$)", format="R$ %,.2f"),
        "Fat. Total": st.column_config.NumberColumn("Fat. Total (R$)", format="R$ %,.2f"),
        "% Total": st.column_config.NumberColumn("% Atingido", format="%.1f%%", disabled=True),
        "UP": st.column_config.NumberColumn("UP (R$)", format="R$ %,.2f", disabled=True),
        "LOSS": st.column_config.NumberColumn("LOSS (R$)", format="R$ %,.2f", disabled=True),
    },
    key="editor_historico"
)

# Recálculo automático de UP / LOSS e % Total
calc_hist = edited_hist.copy()
calc_hist["% Total"] = np.where(calc_hist["Meta Total"] > 0, (calc_hist["Fat. Total"] / calc_hist["Meta Total"]) * 100, 0.0)
calc_hist["UP"] = np.maximum(calc_hist["Fat. Total"] - calc_hist["Meta Total"], 0)
calc_hist["LOSS"] = np.maximum(calc_hist["Meta Total"] - calc_hist["Fat. Total"], 0)

st.markdown("#### Evolução Mensal (Meta vs Faturado)")
st.line_chart(calc_hist.set_index("Mês")[["Meta Total", "Fat. Total"]], color=[BLUE, GREEN])
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 4: PRINCIPAIS CLIENTES =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("4. PRINCIPAIS CLIENTES (TOP 5)")

st.session_state.clientes = st.data_editor(
    st.session_state.clientes,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    key="editor_clientes",
    column_config={
        "RECEITA": st.column_config.NumberColumn("RECEITA (R$)", format="R$ %,.2f"),
        "MÊS ANTERIOR": st.column_config.NumberColumn("MÊS ANTERIOR (R$)", format="R$ %,.2f"),
        "ANO ANTERIOR": st.column_config.NumberColumn("ANO ANTERIOR (R$)", format="R$ %,.2f"),
        "RENTABILIDADE": st.column_config.NumberColumn("RENTABILIDADE %", format="%.1f%%"),
        "SLA": st.column_config.NumberColumn("SLA %", format="%.1f%%"),
        "CONSIDERAÇÕES": st.column_config.TextColumn("CONSIDERAÇÕES / OBS")
    }
)

df_valid_cli = st.session_state.clientes[st.session_state.clientes["CLIENTE"].astype(str).str.strip() != ""]
tot_top5 = df_valid_cli["RECEITA"].sum()

c_tot1, c_tot2 = st.columns(2)
with c_tot1:
    st.metric("TOTAL TOP 5 CLIENTES", f"R$ {fmt_br(tot_top5)}")
with c_tot2:
    st.metric("TOTAL CARTEIRA GERAL", f"R$ {fmt_br(tot_top5)}")
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 5: PIPELINE & UPCOMING =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("5. PIPELINE — NOVOS NEGÓCIOS FECHADOS (MÊS)")

st.session_state.fechados = st.data_editor(
    st.session_state.fechados,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="editor_fechados",
    column_config={
        "PROJEÇÃO MÊS": st.column_config.NumberColumn(format="R$ %,.2f"),
        "FATURADO MÊS": st.column_config.NumberColumn(format="R$ %,.2f"),
    }
)

st.subheader("UPCOMING — FECHADO PARA INICIAR")
st.session_state.upcoming = st.data_editor(
    st.session_state.upcoming,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="editor_upcoming",
    column_config={
        "PROJEÇÃO MÊS": st.column_config.NumberColumn(format="R$ %,.2f"),
    }
)
st.markdown('</div>', unsafe_allow_html=True)

# ===== SEÇÃO 6: EXPORTAÇÃO =====
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("6. EXPORTAR RELATÓRIO")

col_exp1, col_exp2 = st.columns(2)

# Exportar Excel
with col_exp1:
    if st.button("📥 Gerar Excel Completo", use_container_width=True):
        output_xl = BytesIO()
        with pd.ExcelWriter(output_xl, engine="openpyxl") as writer:
            pd.DataFrame([{
                "Nome": nome, "Mês": mes, "Data": data_ref,
                "Meta Gerência": fmt_br(meta), "Faturado Alcançado": fmt_br(faturado),
                "% Alcançado": f"{pct_realizado:.1f}%".replace(".", ","), "Projeção": fmt_br(projecao)
            }]).to_excel(writer, sheet_name="Resumo_Executivo", index=False)
            
            calc_hist_fmt = calc_hist.copy()
            for col in ["Meta Total", "Fat. Total", "UP", "LOSS"]:
                calc_hist_fmt[col] = calc_hist_fmt[col].apply(fmt_br)
            calc_hist_fmt.to_excel(writer, sheet_name="Historico", index=False)
            
            st.session_state.clientes.to_excel(writer, sheet_name="Principais_Clientes", index=False)
            st.session_state.fechados.to_excel(writer, sheet_name="Novos_Negocios", index=False)
            st.session_state.upcoming.to_excel(writer, sheet_name="Upcoming", index=False)
        
        st.download_button(
            label="⬇️ Baixar Planilha (.xlsx)",
            data=output_xl.getvalue(),
            file_name=f"painel_comercial_{mes.lower()}_{nome.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Exportar PowerPoint (.pptx)
with col_exp2:
    if st.button("📊 Gerar Apresentação PowerPoint (.pptx)", use_container_width=True):
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        
        slide1 = prs.slides.add_slide(blank_slide_layout)
        txBox = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tf = txBox.text_frame
        tf.text = f"APRESENTAÇÃO TIME COMERCIAL - {mes.upper()}"
        
        p2 = tf.add_paragraph()
        p2.text = f"Executivo: {nome} | Faturado: R$ {fmt_br(faturado)} / Meta: R$ {fmt_br(meta)} ({pct_realizado:.1f}%)".replace(".", ",")
        p2.font.size = Pt(18)
        p2.font.color.rgb = RGBColor(30, 132, 73)
        
        output_ppt = BytesIO()
        prs.save(output_ppt)
        
        st.download_button(
            label="⬇️ Baixar Apresentação (.pptx)",
            data=output_ppt.getvalue(),
            file_name=f"apresentacao_comercial_{mes.lower()}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

st.markdown('</div>', unsafe_allow_html=True)
