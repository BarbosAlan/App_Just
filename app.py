# ============================================================
# APP WEB - EXTRATO GENIAL | Design Stitch
# Redesign com paleta de cores do Stitch (Azul/Verde/Dourado)
# Rodar com: streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re
import unicodedata
import traceback

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
    page_title="Extrato Genial",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ============================================================
# SISTEMA DE DESIGN - STITCH DESIGN SYSTEM
# Paleta: Azul (#2563eb), Verde (#16a34a), Dourado (#d97706)
# ============================================================

DESIGN_TOKENS = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

        :root {
            /* Stitch / Equity Foundry palette (from stitch/equity_foundry/DESIGN.md) */
            --primary: #003b5a;
            --primary-container: #1a5276;
            --secondary: #006b58;
            --tertiary: #4d3100;
            --error: #ba1a1a;

            /* Surfaces */
            --background: #fcf8fb;
            --surface: #fcf8fb;
            --surface-container-low: #f6f3f5;
            --surface-container-lowest: #ffffff;
            --surface-container-high: #eae7ea;
            --outline: #72787f;
            --outline-variant: #c1c7cf;

            /* Text */
            --on-surface: #1b1b1d;
            --on-surface-variant: #41474e;

            /* Radius */
            --radius-sm: 6px;
            --radius: 10px;
            --radius-lg: 16px;
            --radius-xl: 24px;
            --radius-full: 9999px;

            /* Shadows (use sparingly) */
            --shadow-sm: 0 1px 2px rgba(0,0,0,0.05);
            --shadow-md: 0 10px 32px rgba(0,0,0,0.08);

            /* Transitions */
            --transition-fast: 150ms ease;
            --transition: 250ms ease;
        }

        * {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }

        html, body {
            background: var(--background) !important;
            color: var(--on-surface) !important;
        }

        /* Hide Streamlit chrome */
        #MainMenu, footer, header {visibility: hidden;}

        /* Header */
        .app-header {
            background: rgba(255, 255, 255, 0.92);
            backdrop-filter: blur(18px);
            padding: 1rem 1.25rem;
            border-radius: var(--radius-xl);
            box-shadow: var(--shadow-sm);
            margin: 0 0 1.25rem 0;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .app-header-brand {
            display: flex;
            align-items: center;
            gap: 0.85rem;
        }

        .app-header-icon {
            width: 42px;
            height: 42px;
            background: linear-gradient(145deg, var(--primary), var(--primary-container));
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 1.15rem;
        }

        .app-header-title {
            font-size: 1.125rem;
            font-weight: 800;
            letter-spacing: -0.3px;
            color: var(--primary);
            margin: 0;
            line-height: 1.15;
        }

        .app-header-subtitle {
            font-size: 0.8rem;
            color: var(--on-surface-variant);
            font-weight: 500;
            margin: 0.2rem 0 0 0;
        }

        /* Sections (no hard borders; surface transitions only) */
        .section {
            background: linear-gradient(180deg, var(--surface-container-low) 0%, #f2eff1 100%);
            border-radius: var(--radius-xl);
            padding: 1.75rem;
            margin-bottom: 1.5rem;
        }

        .section-title {
            font-size: 1rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: var(--primary);
            margin: 0 0 1rem 0;
        }

        /* Cards */
        .card {
            background: var(--surface-container-lowest);
            border-radius: var(--radius-lg);
            padding: 1.25rem 1.25rem;
            box-shadow: var(--shadow-sm);
            position: relative;
            overflow: hidden;
        }

        .card-signature::before {
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
            background: var(--primary);
        }

        .card-signature.success::before { background: var(--secondary); }
        .card-signature.warning::before { background: var(--tertiary); }
        .card-signature.error::before { background: var(--error); }

        .card-label {
            font-size: 0.72rem;
            font-weight: 700;
            letter-spacing: 0.16em;
            text-transform: uppercase;
            color: var(--outline);
            margin-top: 0.35rem;
        }

        .card-number {
            font-size: 1.85rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            color: var(--primary);
            font-variant-numeric: tabular-nums;
            margin: 0;
            line-height: 1.1;
        }

        /* Upload zone (glass) */
        .upload-zone {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(24px);
            border: 0.5px dashed rgba(193,199,207,0.9);
            border-radius: var(--radius-xl);
            padding: 2.25rem 1.75rem;
            text-align: center;
            transition: all var(--transition);
        }
        .upload-zone:hover {
            border-color: var(--primary);
            background: rgba(203,230,255,0.14);
            box-shadow: var(--shadow-md);
        }

        /* Buttons */
        .stButton > button {
            background: linear-gradient(145deg, var(--primary), var(--primary-container)) !important;
            color: #fff !important;
            border: none !important;
            border-radius: 14px !important;
            padding: 0.9rem 1.25rem !important;
            font-weight: 700 !important;
            letter-spacing: 0.10em !important;
            text-transform: uppercase !important;
            transition: transform var(--transition-fast), box-shadow var(--transition-fast), opacity var(--transition-fast) !important;
            box-shadow: 0 10px 22px rgba(0, 59, 90, 0.18) !important;
        }
        .stButton > button:hover { transform: translateY(-1px) !important; }
        .stButton > button:disabled { opacity: 0.55 !important; }

        /* Success banner */
        .banner {
            background: rgba(0, 107, 88, 0.10);
            border: 1px solid rgba(0, 107, 88, 0.18);
            border-radius: var(--radius-lg);
            padding: 0.85rem 1rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }
        .banner-dot {
            width: 22px;
            height: 22px;
            border-radius: 999px;
            background: var(--secondary);
            color: #fff;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 900;
            font-size: 0.85rem;
        }

        /* Divider (ghost) */
        .divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(193,199,207,0.45), transparent);
            margin: 1.25rem 0;
        }

        /* Dataframe: softer */
        .stDataFrame, .stDataFrame div, table { border: none !important; }
        thead tr th {
            background: var(--surface-container-high) !important;
            font-size: 0.72rem !important;
            letter-spacing: 0.14em !important;
            text-transform: uppercase !important;
            color: var(--on-surface-variant) !important;
        }
        tbody tr:nth-child(even) td { background: var(--surface-container-low) !important; }
        tbody tr:nth-child(odd) td { background: var(--surface-container-lowest) !important; }

        /* Footer */
        .app-footer {
            text-align: center;
            padding: 1.5rem 0 0.75rem 0;
            color: var(--outline);
            font-size: 0.8rem;
        }

        @media (max-width: 768px) {
            .section { padding: 1.25rem; }
            .card-number { font-size: 1.6rem; }
        }
    </style>
"""

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"\s+", " ", text)
    return text


_TARGETS_NORM = {_normalize_text(t) for t in TARGETS}


def formatar_excel(df_final):
    """Gera o arquivo Excel formatado."""
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
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 22

    for idx, row in df_final.iterrows():
        r = idx + 2
        hist = row["HISTORICO"]

        dc = ws.cell(row=r, column=1, value=row["Data"])
        dc.number_format = "DD/MM/YYYY"
        dc.alignment = center
        dc.border = border

        hc = ws.cell(row=r, column=2, value=hist)
        hc.alignment = left_align
        hc.border = border
        hc.font = Font(name="Arial", size=10)

        vc = ws.cell(row=r, column=3, value=row["Valor"])
        vc.number_format = "#,##0.00"
        vc.alignment = center
        vc.border = border
        if hist in TARGETS:
            cor = "C00000" if row["Valor"] < 0 else "375623"
            vc.font = Font(name="Arial", color=cor, size=10, bold=True)
        else:
            vc.font = Font(name="Arial", size=10)

        hlc = ws.cell(row=r, column=4, value=row["HISTORICO DE LANÇAMENTO"])
        hlc.alignment = left_align
        hlc.border = border
        hlc.font = Font(name="Arial", size=10)

    ws.freeze_panes = "A2"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def processar_arquivo(uploaded_file):
    """Processa o arquivo Excel."""
    try:
        df = pd.read_excel(uploaded_file)

        colunas_obrigatorias = ["Data", "HISTORICO", "Valor", "HISTORICO DE LANÇAMENTO"]
        colunas_faltando = [col for col in colunas_obrigatorias if col not in df.columns]
        if colunas_faltando:
            return None, None, f"❌ Colunas faltando: {', '.join(colunas_faltando)}", None

        df["HISTORICO"] = df["HISTORICO"].astype(str).str.strip()
        df["HISTORICO_NORM"] = df["HISTORICO"].map(_normalize_text)

        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date
        if df["Data"].isna().any():
            invalid = int(df["Data"].isna().sum())
            return None, None, f"❌ A coluna **Data** tem {invalid} valor(es) inválido(s).", None

        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
        if df["Valor"].isna().any():
            invalid = int(df["Valor"].isna().sum())
            return None, None, f"❌ A coluna **Valor** tem {invalid} valor(es) inválido(s).", None

        df_targets = df[df["HISTORICO_NORM"].isin(_TARGETS_NORM)].copy()
        df_others = df[~df["HISTORICO_NORM"].isin(_TARGETS_NORM)].copy()

        # Canonicaliza o histórico dos targets para manter consistência no arquivo final
        norm_to_canonical = {_normalize_text(t): t for t in TARGETS}
        df_targets["HISTORICO"] = df_targets["HISTORICO_NORM"].map(norm_to_canonical).fillna(df_targets["HISTORICO"])

        grouped = df_targets.groupby(["Data", "HISTORICO"], sort=False)["Valor"].sum().reset_index()
        grouped["HISTORICO DE LANÇAMENTO"] = grouped["HISTORICO"]

        df_others_clean = df_others[colunas_obrigatorias].copy()
        df_final = pd.concat([df_others_clean, grouped], ignore_index=True)
        # Ordenação ajuda leitura do “tratado” (mantém comportamento atual do app)
        df_final = df_final.sort_values(["Data", "HISTORICO"], kind="stable").reset_index(drop=True)

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

        return df_final, stats, None, None

    except Exception as e:
        debug = "".join(traceback.format_exception(type(e), e, e.__traceback__))
        return None, None, "❌ Erro ao processar o arquivo.", debug


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
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


def _reset_uploader():
    st.session_state.uploaded_file = None
    st.session_state.uploader_key += 1

# ============================================================
# RENDERIZAÇÃO
# ============================================================

st.markdown(DESIGN_TOKENS, unsafe_allow_html=True)

# HEADER
st.markdown("""
    <div class="app-header">
        <div class="app-header-brand">
            <div class="app-header-icon">💎</div>
            <div>
                <p class="app-header-title">Tratador de Extrato Genial</p>
                <p class="app-header-subtitle">Tratamento automatizado para importação contábil</p>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

# NAVEGAÇÃO
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
    st.button("📁 Histórico", use_container_width=True, disabled=True, type="secondary")

# ============================================================
# PÁGINA: DASHBOARD
# ============================================================

if st.session_state.current_page == "dashboard":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Visão geral</div>', unsafe_allow_html=True)

    if st.session_state.processing_done and st.session_state.stats:
        stats = st.session_state.stats
        has_real_data = True
    else:
        stats = {"total_saida": 0, "total_payin": 0, "total_payout": 0, "saldo": 0}
        has_real_data = False

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div class="card card-signature">
                <div class="card-number">{stats['total_saida'] if has_real_data else "—"}</div>
                <div class="card-label">Transações</div>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="card card-signature success">
                <div class="card-number">{f"R$ {stats['total_payin']:,.0f}" if has_real_data and stats['total_payin'] > 0 else "—"}</div>
                <div class="card-label">Pay In</div>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="card card-signature error">
                <div class="card-number">{f"R$ {abs(stats['total_payout']):,.0f}" if has_real_data and stats['total_payout'] < 0 else "—"}</div>
                <div class="card-label">Pay Out</div>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="card card-signature {'success' if stats.get('saldo', 0) >= 0 else 'error'}">
                <div class="card-number">{f"R$ {stats['saldo']:,.0f}" if has_real_data else "—"}</div>
                <div class="card-label">Saldo</div>
            </div>
        """, unsafe_allow_html=True)

    if not has_real_data:
        st.info("💡 **Nenhum dado ainda.** Vá para **Processamento** para fazer upload do extrato.")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Como funciona</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="card"><div style="font-size: 2rem; margin-bottom: 0.5rem;">1️⃣</div><div style="font-weight: 700; margin-bottom: 0.5rem;">Upload</div><div style="color: var(--text-secondary); font-size: 0.9rem;">Faça upload do arquivo Excel (.xlsx)</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card"><div style="font-size: 2rem; margin-bottom: 0.5rem;">2️⃣</div><div style="font-weight: 700; margin-bottom: 0.5rem;">Processamento</div><div style="color: var(--text-secondary); font-size: 0.9rem;">Agrupamos PAY IN e PAY OUT automaticamente</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="card"><div style="font-size: 2rem; margin-bottom: 0.5rem;">3️⃣</div><div style="font-weight: 700; margin-bottom: 0.5rem;">Download</div><div style="color: var(--text-secondary); font-size: 0.9rem;">Baixe o arquivo tratado e formatado</div></div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# PÁGINA: PROCESSAMENTO
# ============================================================

elif st.session_state.current_page == "processamento":
    st.markdown('<div class="section">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Processamento</div>', unsafe_allow_html=True)

    if not st.session_state.processing_done:
        # Estado 1: Upload
        st.markdown('<div class="upload-zone">', unsafe_allow_html=True)
        uploaded_file = st.file_uploader(
            "Selecione o arquivo Excel (.xlsx)",
            type=["xlsx"],
            help="Colunas: Data, HISTORICO, Valor, HISTORICO DE LANÇAMENTO",
            label_visibility="collapsed",
            key=f"uploader_{st.session_state.uploader_key}",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_file:
            st.session_state.uploaded_file = uploaded_file

        if st.session_state.uploaded_file is not None:
            f = st.session_state.uploaded_file
            size_mb = None
            try:
                size_mb = (getattr(f, "size", None) or 0) / (1024 * 1024)
            except Exception:
                size_mb = None

            left, right = st.columns([10, 2])
            with left:
                st.markdown(
                    f"""
                    <div class="card card-signature success" style="display:flex; align-items:center; justify-content:space-between; gap:1rem;">
                        <div>
                            <div style="font-weight:800; color: var(--on-surface); font-size:0.95rem;">{f.name}</div>
                            <div class="card-label" style="margin-top:0.4rem;">
                                Planilha Excel{f" • {size_mb:.1f} MB" if size_mb and size_mb > 0 else ""}
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            with right:
                if st.button("✕", help="Remover arquivo", use_container_width=True):
                    _reset_uploader()
                    st.rerun()

            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Bento info grid (estilo do mock)
        col_i1, col_i2, col_i3 = st.columns(3)
        with col_i1:
            st.markdown(
                """
                <div class="card">
                    <div style="font-weight:800; color: var(--primary); letter-spacing:0.14em; text-transform:uppercase; font-size:0.72rem;">Privacidade</div>
                    <div style="margin-top:0.5rem; color: var(--on-surface-variant); font-size:0.9rem;">
                        Seus dados são processados localmente e nunca armazenados.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_i2:
            st.markdown(
                """
                <div class="card">
                    <div style="font-weight:800; color: var(--primary); letter-spacing:0.14em; text-transform:uppercase; font-size:0.72rem;">Velocidade</div>
                    <div style="margin-top:0.5rem; color: var(--on-surface-variant); font-size:0.9rem;">
                        Processamento de milhares de linhas em poucos segundos.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col_i3:
            st.markdown(
                """
                <div class="card">
                    <div style="font-weight:800; color: var(--primary); letter-spacing:0.14em; text-transform:uppercase; font-size:0.72rem;">Padronização</div>
                    <div style="margin-top:0.5rem; color: var(--on-surface-variant); font-size:0.9rem;">
                        Conversão automática para formatos contábeis aceitos.
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            can_process = st.session_state.uploaded_file is not None
            processar_btn = st.button("⚡ PROCESSAR ARQUIVO", type="primary", disabled=not can_process, use_container_width=True)

        if processar_btn and st.session_state.uploaded_file is not None:
            with st.spinner("⏳ Processando..."):
                df_final, stats, error, debug = processar_arquivo(st.session_state.uploaded_file)

            if error:
                st.error(error)
                if debug:
                    with st.expander("Detalhes técnicos (para suporte)"):
                        st.code(debug)
            else:
                st.session_state.df_final = df_final
                st.session_state.stats = stats
                st.session_state.processing_done = True
                st.rerun()

    else:
        # Estado 2: Sucesso
        stats = st.session_state.stats
        df_final = st.session_state.df_final

        st.markdown("""
            <div class="banner">
                <div class="banner-dot">✓</div>
                <div>
                    <div style="font-weight: 700; color: var(--secondary); font-size: 0.9rem;">Arquivo processado com sucesso</div>
                    <div style="color: var(--on-surface-variant); font-size: 0.8rem;">O tratamento foi concluído seguindo as regras de agrupamento.</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Cards de Estatística
        st.markdown("### Resumo")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="card card-signature"><div class="card-number">{stats["linhas_originais"]}</div><div class="card-label">Originais</div></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="card card-signature success"><div class="card-number">{stats["agrupados"]}</div><div class="card-label">Agrupados</div></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="card card-signature warning"><div class="card-number">{stats["outros_mantidos"]}</div><div class="card-label">Outros</div></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="card card-signature"><div class="card-number">{stats["total_saida"]}</div><div class="card-label">Total</div></div>', unsafe_allow_html=True)

        # Métricas financeiras
        col5, col6, col7 = st.columns(3)
        with col5:
            st.markdown(f'<div class="card card-signature success"><div class="card-number" style="font-size: 1.8rem;">R$ {stats["total_payin"]:,.2f}</div><div class="card-label">Pay In</div></div>', unsafe_allow_html=True)
        with col6:
            st.markdown(f'<div class="card card-signature error"><div class="card-number" style="font-size: 1.8rem; background: linear-gradient(135deg, var(--danger), #ef4444); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">R$ {abs(stats["total_payout"]):,.2f}</div><div class="card-label">Pay Out</div></div>', unsafe_allow_html=True)
        with col7:
            saldo_cls = 'success' if stats['saldo'] >= 0 else 'error'
            st.markdown(f'<div class="card card-signature {saldo_cls}"><div class="card-number" style="font-size: 1.8rem;">R$ {stats["saldo"]:,.2f}</div><div class="card-label">Saldo</div></div>', unsafe_allow_html=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Preview
        st.markdown("### Pré-visualização")
        st.dataframe(df_final.head(20), use_container_width=True, hide_index=True)
        if len(df_final) > 20:
            st.caption(f"Mostrando 20 de {len(df_final)} linhas")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        # Download
        st.markdown("### Download")
        excel_output = formatar_excel(df_final)
        nome_arquivo = st.session_state.uploaded_file.name.replace(".xlsx", "_tratado.xlsx")

        st.download_button(
            label="📥 BAIXAR EXCEL TRATADO",
            data=excel_output,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        if st.button("🔄 NOVO ARQUIVO", type="secondary", use_container_width=True):
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
        <p>💎 Extrato Genial • Automatização Inteligente</p>
    </div>
""", unsafe_allow_html=True)
