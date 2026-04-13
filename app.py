# ============================================================
# APP WEB - TRATADOR DE EXTRATO GENIAL
# Rodar com: streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import date

# Configuração da página
st.set_page_config(
    page_title="Tratador de Extrato Genial",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

TARGET_PAYOUT = "PAGAMENTO DE PAY OUT - PARCEIRO"
TARGET_PAYIN  = "RECEBIMENTO DE PAY IN DE PARCEIRO"
TARGETS       = [TARGET_PAYOUT, TARGET_PAYIN]


def formatar_excel(df_final):
    """Gera o arquivo Excel formatado igual ao script original."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Extrato Tratado"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", start_color="1F4E79")
    payin_fill   = PatternFill("solid", start_color="E2EFDA")
    payout_fill  = PatternFill("solid", start_color="FDECEA")
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

        if hist == TARGET_PAYIN:
            fill = payin_fill
        elif hist == TARGET_PAYOUT:
            fill = payout_fill
        else:
            fill = None

        dc                = ws.cell(row=r, column=1, value=row["Data"])
        dc.number_format  = "DD/MM/YYYY"
        dc.alignment      = center
        dc.border         = border
        if fill: dc.fill  = fill

        hc           = ws.cell(row=r, column=2, value=hist)
        hc.alignment = left_align
        hc.border    = border
        hc.font      = Font(name="Arial", size=10)
        if fill: hc.fill = fill

        vc                = ws.cell(row=r, column=3, value=row["Valor"])
        vc.number_format  = "#,##0.00"
        vc.alignment      = center
        vc.border         = border
        if hist in TARGETS:
            cor       = "C00000" if row["Valor"] < 0 else "375623"
            vc.font   = Font(name="Arial", color=cor, size=10, bold=True)
        else:
            vc.font   = Font(name="Arial", size=10)
        if fill: vc.fill = fill

        hlc           = ws.cell(row=r, column=4, value=row["HISTORICO DE LANÇAMENTO"])
        hlc.alignment = left_align
        hlc.border    = border
        hlc.font      = Font(name="Arial", size=10)
        if fill: hlc.fill = fill

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
            st.error(f"❌ Colunas faltando no arquivo: {', '.join(colunas_faltando)}")
            return None, None

        df["HISTORICO"] = df["HISTORICO"].str.strip()
        df["Data"] = pd.to_datetime(df["Data"]).dt.date

        df_targets = df[df["HISTORICO"].isin(TARGETS)].copy()
        df_others  = df[~df["HISTORICO"].isin(TARGETS)].copy()

        grouped = df_targets.groupby(["Data", "HISTORICO"], sort=False)["Valor"].sum().reset_index()
        grouped["HISTORICO DE LANÇAMENTO"] = grouped["HISTORICO"]

        df_others_clean = df_others[colunas_obrigatorias].copy()
        df_final = pd.concat([df_others_clean, grouped], ignore_index=True)
        df_final = df_final.sort_values(["Data", "HISTORICO"]).reset_index(drop=True)

        stats = {
            "linhas_originais": len(df),
            "outros_mantidos": len(df_others_clean),
            "agrupados": len(grouped),
            "total_saida": len(df_final),
        }

        return df_final, stats

    except Exception as e:
        st.error(f"❌ Erro ao processar arquivo: {str(e)}")
        return None, None


# ============================================================
# INTERFACE DO APP
# ============================================================

# CSS customizado
st.markdown("""
    <style>
        .main-header {
            font-size: 2.5rem;
            font-weight: 700;
            color: #1F4E79;
            text-align: center;
            margin-bottom: 0.5rem;
        }
        .sub-header {
            font-size: 1.1rem;
            color: #666;
            text-align: center;
            margin-bottom: 2rem;
        }
        .stButton button {
            background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
            color: white;
            border: none;
            border-radius: 8px;
            padding: 12px 32px;
            font-size: 1.1rem;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        .stButton button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(31,78,121,0.3);
        }
        .stButton button:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }
        .stat-card {
            background: white;
            padding: 1.5rem;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            text-align: center;
            border-left: 4px solid #1F4E79;
        }
        .stat-number {
            font-size: 2rem;
            font-weight: 700;
            color: #1F4E79;
        }
        .stat-label {
            font-size: 0.9rem;
            color: #666;
            margin-top: 0.5rem;
        }
        .upload-box {
            border: 2px dashed #1F4E79;
            border-radius: 12px;
            padding: 2rem;
            text-align: center;
            background: #F8F9FA;
        }
        .divider {
            height: 2px;
            background: linear-gradient(90deg, transparent, #1F4E79, transparent);
            margin: 2rem 0;
        }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 Tratador de Extrato Genial</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Faça upload do arquivo Excel, processe e baixe o arquivo formatado automaticamente</div>', unsafe_allow_html=True)

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# Upload
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "📁 Selecione o arquivo Excel (.xlsx)",
        type=["xlsx"],
        help="Selecione o arquivo Excel com o extrato para processar",
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

# Botão processar
col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    processar_btn = st.button(
        "⚡ PROCESSAR ARQUIVO",
        type="primary",
        disabled=uploaded_file is None,
        use_container_width=True
    )

# Processamento
if processar_btn and uploaded_file is not None:
    with st.spinner("⏳ Processando arquivo..."):
        df_final, stats = processar_arquivo(uploaded_file)

    if df_final is not None and stats is not None:
        st.success("✅ Processamento concluído com sucesso!")

        # Stats em cards
        st.markdown("### 📈 Resumo do Processamento")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-number">{stats['linhas_originais']}</div>
                    <div class="stat-label">Linhas Originais</div>
                </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-number">{stats['outros_mantidos']}</div>
                    <div class="stat-label">Outros Mantidos</div>
                </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-number">{stats['agrupados']}</div>
                    <div class="stat-label">PAY IN/OUT Agrupados</div>
                </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-number">{stats['total_saida']}</div>
                    <div class="stat-label">Total no Arquivo</div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("")

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

        # Download
        st.markdown("")
        st.markdown("### 📥 Download do Arquivo Tratado")

        excel_output = formatar_excel(df_final)
        nome_arquivo = uploaded_file.name.replace(".xlsx", "_tratado.xlsx")

        st.download_button(
            label="📥 BAIXAR ARQUIVO EXCEL TRATADO",
            data=excel_output,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# Footer
st.markdown(
    """
    <div style="text-align: center; color: #999; font-size: 0.85rem; padding: 1rem 0;">
        <p>📊 Tratador de Extrato Genial • Automatização do processamento de extratos</p>
    </div>
    """,
    unsafe_allow_html=True
)
