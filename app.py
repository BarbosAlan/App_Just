# ============================================================
# APP WEB - EXTRATO GENIAL | "The Precise Curator"
# Redesign - Identidade Visual Premium
# Rodar com: streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import date
import random

# ============================================================
# CONSTANTES
# ============================================================

TARGET_PAYOUT = "PAGAMENTO DE PAY OUT - PARCEIRO"
TARGET_PAYIN  = "RECEBIMENTO DE PAY IN DE PARCEIRO"
TARGET_BLOQUEIO = "BLOQUEIO - DEPOSITO JUDICIAL"
TARGET_DESBLOQUEIO = "DESBLOQUEIO JUDICIAL"
TARGETS = [TARGET_PAYOUT, TARGET_PAYIN, TARGET_BLOQUEIO, TARGET_DESBLOQUEIO]

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Extrato Genial — The Precise Curator",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# SISTEMA DE DESIGN - "THE PRECISE CURATOR"
# ============================================================

DESIGN_TOKENS = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        :root {
            /* Paleta Primária */
            --primary-foundation: #003b5a;
            --primary-container: #1a5276;
            --primary-light: #2E75B6;

            /* Paleta Secundária */
            --secondary: #006b58;
            --secondary-light: #008c73;

            /* Superfícies */
            --surface-base: #fcf8fb;
            --surface-section: #f6f3f5;
            --surface-card: #ffffff;

            /* Texto */
            --text-primary: #1a1a2e;
            --text-secondary: #6b6b7b;
            --text-muted: #9b9bab;

            /* Estado */
            --success: #006b58;
            --error: #c0392b;
            --warning: #f39c12;

            /* Sombras */
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.06);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
            --shadow-lg: 0 8px 24px rgba(0,0,0,0.12);

            /* Bordas */
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 16px;
        }

        * {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        /* Remover linhas de tabela padrão */
        .stDataFrame, .stDataFrame div, .stDataFrame table {
            border: none !important;
        }

        table {
            border-collapse: collapse !important;
        }

        thead tr, tbody tr {
            border: none !important;
        }

        th {
            border-bottom: 2px solid #e8e8ec !important;
            background: var(--surface-section) !important;
        }

        td {
            border-bottom: 1px solid #f0f0f4 !important;
        }

        /* Esconder bordas padrão do Streamlit */
        .stTextInput > div > div, .stSelectbox > div > div {
            border: none !important;
        }

        /* HEADER */
        .app-header {
            background: var(--surface-card);
            padding: 1rem 2rem;
            box-shadow: var(--shadow-sm);
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-radius: var(--radius-md);
            margin-bottom: 2rem;
        }

        .app-header-brand {
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .app-header-icon {
            width: 48px;
            height: 48px;
            background: linear-gradient(135deg, var(--primary-foundation), var(--primary-light));
            border-radius: var(--radius-sm);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
        }

        .app-header-title {
            font-size: 1.5rem;
            font-weight: 700;
            color: var(--primary-foundation);
            margin: 0;
            letter-spacing: -0.02em;
        }

        .app-header-subtitle {
            font-size: 0.8rem;
            color: var(--text-secondary);
            font-weight: 400;
            margin: 0;
        }

        /* NAVEGAÇÃO */
        .nav-bar {
            display: flex;
            gap: 0.5rem;
            background: var(--surface-section);
            padding: 0.5rem;
            border-radius: var(--radius-md);
            margin-bottom: 2rem;
        }

        .nav-item {
            padding: 0.75rem 1.5rem;
            border-radius: var(--radius-sm);
            cursor: pointer;
            transition: all 0.2s ease;
            font-weight: 500;
            font-size: 0.9rem;
            color: var(--text-secondary);
            text-align: center;
            flex: 1;
        }

        .nav-item:hover {
            background: var(--surface-card);
            color: var(--text-primary);
        }

        .nav-item.active {
            background: var(--surface-card);
            color: var(--primary-foundation);
            box-shadow: var(--shadow-sm);
            font-weight: 600;
        }

        .nav-item.disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        /* CARDS */
        .card {
            background: var(--surface-card);
            border-radius: var(--radius-md);
            padding: 1.5rem;
            box-shadow: var(--shadow-sm);
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }

        .card:hover {
            box-shadow: var(--shadow-md);
            transform: translateY(-2px);
        }

        /* SIGNATURE STYLE - Barra lateral de cor */
        .card-signature {
            border-left: none;
        }

        .card-signature::before {
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            height: 100%;
            width: 4px;
            background: linear-gradient(180deg, var(--primary-foundation), var(--primary-light));
            border-radius: 4px 0 0 4px;
        }

        .card-signature.success::before {
            background: linear-gradient(180deg, var(--secondary), var(--secondary-light));
        }

        .card-signature.warning::before {
            background: linear-gradient(180deg, var(--warning), #f5b041);
        }

        .card-signature.error::before {
            background: linear-gradient(180deg, var(--error), #e74c3c);
        }

        .card-number {
            font-size: 2.5rem;
            font-weight: 800;
            color: var(--primary-foundation);
            line-height: 1;
            margin-bottom: 0.5rem;
            letter-spacing: -0.03em;
        }

        .card-label {
            font-size: 0.85rem;
            color: var(--text-secondary);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        /* SEÇÕES */
        .section {
            background: var(--surface-section);
            border-radius: var(--radius-lg);
            padding: 2rem;
            margin-bottom: 2rem;
        }

        .section-title {
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--primary-foundation);
            margin-bottom: 1.5rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .section-title-icon {
            font-size: 1.5rem;
        }

        /* UPLOAD BOX */
        .upload-zone {
            border: 2px dashed var(--primary-light);
            border-radius: var(--radius-md);
            padding: 3rem 2rem;
            text-align: center;
            background: var(--surface-card);
            transition: all 0.3s ease;
        }

        .upload-zone:hover {
            border-color: var(--primary-foundation);
            background: var(--surface-section);
        }

        .upload-icon {
            font-size: 3rem;
            margin-bottom: 1rem;
        }

        .upload-text {
            font-size: 1.1rem;
            font-weight: 600;
            color: var(--text-primary);
            margin-bottom: 0.5rem;
        }

        .upload-subtext {
            font-size: 0.9rem;
            color: var(--text-secondary);
        }

        /* BOTÕES */
        .stButton > button {
            background: linear-gradient(135deg, var(--primary-foundation) 0%, var(--primary-container) 100%) !important;
            color: white !important;
            border: none !important;
            border-radius: var(--radius-sm) !important;
            padding: 0.875rem 2rem !important;
            font-size: 1rem !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
            box-shadow: var(--shadow-sm) !important;
        }

        .stButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: var(--shadow-md) !important;
        }

        .stButton > button:disabled {
            opacity: 0.5 !important;
            cursor: not-allowed !important;
            transform: none !important;
        }

        /* STATUS */
        .status-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            padding: 0.5rem 1rem;
            border-radius: 2rem;
            font-size: 0.85rem;
            font-weight: 600;
            background: var(--surface-card);
            box-shadow: var(--shadow-sm);
        }

        .status-badge.success {
            color: var(--success);
            background: rgba(0, 107, 88, 0.1);
        }

        .status-badge.processing {
            color: var(--warning);
            background: rgba(243, 156, 18, 0.1);
        }

        /* DIVIDER */
        .divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, #e0e0e5, transparent);
            margin: 2rem 0;
        }

        /* FOOTER */
        .app-footer {
            text-align: center;
            padding: 2rem 0;
            color: var(--text-muted);
            font-size: 0.8rem;
            font-weight: 400;
        }

        /* ESconder elementos padrão do Streamlit */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Responsividade */
        @media (max-width: 768px) {
            .app-header {
                flex-direction: column;
                gap: 1rem;
                padding: 1rem;
            }

            .nav-bar {
                flex-direction: column;
            }

            .nav-item {
                padding: 0.75rem 1rem;
            }

            .card-number {
                font-size: 2rem;
            }
        }
    </style>
"""

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def formatar_excel(df_final):
    """Gera o arquivo Excel formatado igual ao script original."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato Tratado"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", start_color="1F4E79")
    border       = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )
    center      = Alignment(horizontal="center", vertical="center")
    left_align  = Alignment(horizontal="left", vertical="center")

    headers    = ["Data", "HISTORICO", "Valor", "HISTORICO DE LANÇAMENTO"]
    col_widths = [15, 45, 18, 45]

    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell            = ws.cell(row=1, column=i, value=h)
        cell.font       = header_font
        cell.fill       = header_fill
        cell.alignment  = center
        cell.border     = border
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22

    for idx, row in df_final.iterrows():
        r    = idx + 2
        hist = row["HISTORICO"]

        dc                = ws.cell(row=r, column=1, value=row["Data"])
        dc.number_format  = "DD/MM/YYYY"
        dc.alignment      = center
        dc.border         = border

        hc           = ws.cell(row=r, column=2, value=hist)
        hc.alignment = left_align
        hc.border    = border
        hc.font      = Font(name="Arial", size=10)

        vc                = ws.cell(row=r, column=3, value=row["Valor"])
        vc.number_format  = "#,##0.00"
        vc.alignment      = center
        vc.border         = border
        if hist in TARGETS:
            cor       = "C00000" if row["Valor"] < 0 else "375623"
            vc.font   = Font(name="Arial", color=cor, size=10, bold=True)
        else:
            vc.font   = Font(name="Arial", size=10)

        hlc           = ws.cell(row=r, column=4, value=row["HISTORICO DE LANÇAMENTO"])
        hlc.alignment = left_align
        hlc.border    = border
        hlc.font      = Font(name="Arial", size=10)

    ws.freeze_panes = "A2"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def processar_arquivo(uploaded_file):
    """Processa o arquivo Excel e retorna os resultados."""
    try:
        df = pd.read_excel(uploaded_file)

        # Valida colunas obrigatórias
        colunas_obrigatorias = ["Data", "HISTORICO", "Valor", "HISTORICO DE LANÇAMENTO"]
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        if colunas_faltando:
            return None, None, f"❌ Colunas faltando no arquivo: {', '.join(colunas_faltando)}"

        df["HISTORICO"] = df["HISTORICO"].str.strip()
        df["Data"] = pd.to_datetime(df["Data"]).dt.date

        df_targets = df[df["HISTORICO"].isin(TARGETS)].copy()
        df_others  = df[~df["HISTORICO"].isin(TARGETS)].copy()

        grouped = df_targets.groupby(["Data", "HISTORICO"], sort=False)["Valor"].sum().reset_index()
        grouped["HISTORICO DE LANÇAMENTO"] = grouped["HISTORICO"]

        df_others_clean = df_others[colunas_obrigatorias].copy()
        df_final = pd.concat([df_others_clean, grouped], ignore_index=True)
        df_final = df_final.sort_values(["Data", "HISTORICO"]).reset_index(drop=True)

        # Calcula métricas adicionais
        total_payin = df_final[df_final["HISTORICO"] == TARGET_PAYIN]["Valor"].sum()
        total_payout = df_final[df_final["HISTORICO"] == TARGET_PAYOUT]["Valor"].sum()
        saldo = total_payin + total_payout

        stats = {
            "linhas_originais": len(df),
            "outros_mantidos": len(df_others_clean),
            "agrupados": len(grouped),
            "total_saida": len(df_final),
            "total_payin": total_payin,
            "total_payout": total_payout,
            "saldo": saldo,
        }

        return df_final, stats, None

    except Exception as e:
        return None, None, "❌ Ocorreu um erro ao processar o arquivo. Verifique o formato e tente novamente."


# ============================================================
# ESTADO DA APLICAÇÃO
# ============================================================

if "current_page" not in st.session_state:
    st.session_state.current_page = "dashboard"
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "df_final" not in st.session_state:
    st.session_state.df_final = None
if "stats" not in st.session_state:
    st.session_state.stats = None
if "processing_done" not in st.session_state:
    st.session_state.processing_done = False

# ============================================================
# RENDERIZAÇÃO DA UI
# ============================================================

# Injeta CSS do sistema de design
st.markdown(DESIGN_TOKENS, unsafe_allow_html=True)

# HEADER
st.markdown("""
    <div class="app-header">
        <div class="app-header-brand">
            <div class="app-header-icon">💎</div>
            <div>
                <p class="app-header-title">Extrato Genial</p>
                <p class="app-header-subtitle">The Precise Curator</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# NAVEGAÇÃO
nav_items = [
    {"id": "dashboard", "label": "📊 Dashboard", "disabled": False},
    {"id": "processamento", "label": "⚡ Processamento", "disabled": False},
    {"id": "historico", "label": "📁 Histórico", "disabled": True},
]

st.markdown('<div class="nav-bar">', unsafe_allow_html=True)
for item in nav_items:
    active_class = "active" if st.session_state.current_page == item["id"] else ""
    disabled_class = "disabled" if item["disabled"] else ""
    if item["disabled"]:
        st.markdown(f'<div class="nav-item {disabled_class}">{item["label"]} <span style="font-size: 0.7rem;">(Em breve)</span></div>', unsafe_allow_html=True)
    else:
        if st.markdown(f'<div class="nav-item {active_class}" onclick="window.location.href=\'#{item["id"]}\'">{item["label"]}</div>', unsafe_allow_html=True):
            st.session_state.current_page = item["id"]
st.markdown('</div>', unsafe_allow_html=True)

# Botões de navegação funcionais (Streamlit way)
col_nav1, col_nav2, col_nav3 = st.columns(3)
with col_nav1:
    if st.button("📊 Dashboard", use_container_width=True, type="primary" if st.session_state.current_page == "dashboard" else "secondary"):
        st.session_state.current_page = "dashboard"
        st.rerun()
with col_nav2:
    if st.button("⚡ Processamento", use_container_width=True, type="primary" if st.session_state.current_page == "processamento" else "secondary"):
        st.session_state.current_page = "processamento"
        st.rerun()
with col_nav3:
    if st.button("📁 Histórico", use_container_width=True, disabled=True, type="secondary"):
        pass

# ============================================================
# PÁGINA: DASHBOARD
# ============================================================

if st.session_state.current_page == "dashboard":
    st.markdown("""
        <div class="section">
            <div class="section-title">
                <span class="section-title-icon">📈</span>
                Visão Geral
            </div>
    """, unsafe_allow_html=True)

    # Stats simulados ou baseados em dados processados
    if st.session_state.processing_done and st.session_state.stats:
        stats = st.session_state.stats
        has_real_data = True
    else:
        # Dados simulados para demonstração
        stats = {
            "linhas_originais": 0,
            "outros_mantidos": 0,
            "agrupados": 0,
            "total_saida": 0,
            "total_payin": 0,
            "total_payout": 0,
            "saldo": 0,
        }
        has_real_data = False

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div class="card card-signature">
                <div class="card-number">{stats['total_saida'] if has_real_data else "—"}</div>
                <div class="card-label">Transações Processadas</div>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="card card-signature success">
                <div class="card-number" style="color: var(--secondary);">
                    {f"R$ {stats['total_payin']:,.2f}" if has_real_data and stats['total_payin'] > 0 else "—"}
                </div>
                <div class="card-label">Total Pay In</div>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="card card-signature error">
                <div class="card-number" style="color: var(--error);">
                    {f"R$ {stats['total_payout']:,.2f}" if has_real_data and stats['total_payout'] > 0 else "—"}
                </div>
                <div class="card-label">Total Pay Out</div>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        saldo_color = "var(--secondary)" if stats.get('saldo', 0) >= 0 else "var(--error)"
        st.markdown(f"""
            <div class="card card-signature {'success' if stats.get('saldo', 0) >= 0 else 'error'}">
                <div class="card-number" style="color: {saldo_color};">
                    {f"R$ {stats['saldo']:,.2f}" if has_real_data else "—"}
                </div>
                <div class="card-label">Saldo</div>
            </div>
        """, unsafe_allow_html=True)

    if not has_real_data:
        st.info("💡 **Nenhum dado disponível ainda.** Vá para a aba **Processamento** para fazer upload do seu extrato e visualizar as estatísticas reais.")

    # Seção de informações adicionais
    st.markdown("""
        <div class="divider"></div>
        <div class="section-title" style="margin-top: 2rem;">
            <span class="section-title-icon">ℹ️</span>
            Como Funciona
        </div>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1.5rem;">
            <div class="card">
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">1️⃣</div>
                <div style="font-weight: 600; margin-bottom: 0.5rem;">Upload</div>
                <div style="color: var(--text-secondary); font-size: 0.9rem;">Faça upload do arquivo Excel (.xlsx) com seu extrato</div>
            </div>
            <div class="card">
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">2️⃣</div>
                <div style="font-weight: 600; margin-bottom: 0.5rem;">Processamento</div>
                <div style="color: var(--text-secondary); font-size: 0.9rem;">Agrupamos PAY IN e PAY OUT por dia automaticamente</div>
            </div>
            <div class="card">
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">3️⃣</div>
                <div style="font-weight: 600; margin-bottom: 0.5rem;">Download</div>
                <div style="color: var(--text-secondary); font-size: 0.9rem;">Baixe o arquivo tratado e formatado com cores e bordas</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# PÁGINA: PROCESSAMENTO
# ============================================================

elif st.session_state.current_page == "processamento":
    st.markdown("""
        <div class="section">
            <div class="section-title">
                <span class="section-title-icon">⚡</span>
                Processamento de Extrato
            </div>
    """, unsafe_allow_html=True)

    # Estado 1: Upload
    if not st.session_state.processing_done:
        st.markdown('<div class="upload-zone">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Selecione o arquivo Excel (.xlsx)",
            type=["xlsx"],
            help="O arquivo deve conter as colunas: Data, HISTORICO, Valor, HISTORICO DE LANÇAMENTO",
            label_visibility="collapsed"
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success(f"✅ Arquivo selecionado: **{uploaded_file.name}**")

        # Botão processar
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            processar_btn = st.button(
                "⚡ PROCESSAR ARQUIVO",
                type="primary",
                disabled=uploaded_file is None,
                use_container_width=True
            )

        if processar_btn and uploaded_file is not None:
            with st.spinner("⏳ Processando arquivo..."):
                df_final, stats, error = processar_arquivo(uploaded_file)

            if error:
                st.error(error)
            else:
                st.session_state.df_final = df_final
                st.session_state.stats = stats
                st.session_state.processing_done = True
                st.rerun()

    # Estado 2: Sucesso e Pré-visualização
    else:
        stats = st.session_state.stats
        df_final = st.session_state.df_final

        # Badge de sucesso
        st.markdown('<div class="status-badge success">✅ Processamento concluído com sucesso!</div>', unsafe_allow_html=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Cards de Estatística (Signature Style)
        st.markdown("### 📊 Resumo do Processamento")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
                <div class="card card-signature">
                    <div class="card-number">{stats['linhas_originais']}</div>
                    <div class="card-label">Linhas Originais</div>
                </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
                <div class="card card-signature success">
                    <div class="card-number" style="color: var(--secondary);">{stats['agrupados']}</div>
                    <div class="card-label">PAY IN/OUT Agrupados</div>
                </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
                <div class="card card-signature warning">
                    <div class="card-number" style="color: var(--warning);">{stats['outros_mantidos']}</div>
                    <div class="card-label">Outros Mantidos</div>
                </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
                <div class="card card-signature">
                    <div class="card-number">{stats['total_saida']}</div>
                    <div class="card-label">Total no Arquivo</div>
                </div>
            """, unsafe_allow_html=True)

        # Métricas financeiras adicionais
        col5, col6, col7 = st.columns(3)

        with col5:
            st.markdown(f"""
                <div class="card card-signature success">
                    <div class="card-number" style="color: var(--secondary); font-size: 1.8rem;">
                        R$ {stats['total_payin']:,.2f}
                    </div>
                    <div class="card-label">Total Pay In</div>
                </div>
            """, unsafe_allow_html=True)

        with col6:
            st.markdown(f"""
                <div class="card card-signature error">
                    <div class="card-number" style="color: var(--error); font-size: 1.8rem;">
                        R$ {abs(stats['total_payout']):,.2f}
                    </div>
                    <div class="card-label">Total Pay Out</div>
                </div>
            """, unsafe_allow_html=True)

        with col7:
            saldo_color = "var(--secondary)" if stats['saldo'] >= 0 else "var(--error)"
            st.markdown(f"""
                <div class="card card-signature {'success' if stats['saldo'] >= 0 else 'error'}">
                    <div class="card-number" style="color: {saldo_color}; font-size: 1.8rem;">
                        R$ {stats['saldo']:,.2f}
                    </div>
                    <div class="card-label">Saldo</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Preview dos dados
        st.markdown("### 👀 Pré-visualização dos Dados")
        st.dataframe(
            df_final.head(20),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Data": st.column_config.DateColumn("Data"),
                "HISTORICO": st.column_config.TextColumn("Histórico"),
                "Valor": st.column_config.NumberColumn("Valor", format="%.2f"),
                "HISTORICO DE LANÇAMENTO": st.column_config.TextColumn("Hist. Lançamento"),
            }
        )
        if len(df_final) > 20:
            st.caption(f"Mostrando 20 de {len(df_final)} linhas")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Estado 3: Download
        st.markdown("### 📥 Download do Arquivo Tratado")

        excel_output = formatar_excel(df_final)
        nome_arquivo = st.session_state.uploaded_file.name.replace(".xlsx", "_tratado.xlsx")

        st.download_button(
            label="📥 BAIXAR ARQUIVO EXCEL TRATADO",
            data=excel_output,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        # Botão para novo processamento
        st.markdown("")
        if st.button("🔄 PROCESSAR NOVO ARQUIVO", type="secondary", use_container_width=True):
            st.session_state.processing_done = False
            st.session_state.uploaded_file = None
            st.session_state.df_final = None
            st.session_state.stats = None
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# FOOTER
# ============================================================

st.markdown("""
    <div class="app-footer">
        <div class="divider"></div>
        <p>💎 Extrato Genial • The Precise Curator • Automatização inteligente de extratos</p>
    </div>
""", unsafe_allow_html=True)
