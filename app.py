from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="SCM Analytics Hub",
    layout="wide",
    initial_sidebar_state="expanded"
)

LOGO_PATH = "assets/impress_logo.png"

# Logo na sidebar, se existir
if Path(LOGO_PATH).is_file():
    st.logo(LOGO_PATH, size="large")

# CSS da sidebar + CSS da página inicial
st.markdown(
    """
    <style>

    /* =========================
       SIDEBAR NAVIGATION
       ========================= */

    /* Esconde somente o texto original da navegação lateral */
    [data-testid="stSidebarNav"] ul li a span {
        display: none !important;
    }

    /* Ajusta aparência dos links da sidebar */
    [data-testid="stSidebarNav"] ul li a {
        display: flex !important;
        align-items: center !important;
        min-height: 38px !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        text-decoration: none !important;
    }

    /* Home / app.py */
    [data-testid="stSidebarNav"] ul li:nth-child(1) a::after {
        content: "🏠 Inicio";
        font-size: 14px;
        font-weight: 600;
    }

    /* Cockpit Papel */
    [data-testid="stSidebarNav"] ul li:nth-child(2) a::after {
        content: "🔎 Cockpit Papel";
        font-size: 14px;
        font-weight: 600;
    }

    /* Paper Base */
    [data-testid="stSidebarNav"] ul li:nth-child(3) a::after {
        content: "📈 Paper Base";
        font-size: 14px;
        font-weight: 600;
    }


    /* =========================
       HOME PAGE
       ========================= */

    .main-title {
        font-size: 42px;
        font-weight: 800;
        color: #111827;
        margin-bottom: 6px;
        line-height: 1.15;
    }

    .subtitle {
        font-size: 18px;
        color: #4B5563;
        margin-bottom: 30px;
        line-height: 1.45;
    }

    .hub-card {
        border: 1px solid #E5E7EB;
        border-radius: 18px;
        padding: 26px;
        background: #FFFFFF;
        box-shadow: 0 4px 14px rgba(0,0,0,0.05);
        min-height: 230px;
        transition: all 0.2s ease-in-out;
    }

    .hub-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 22px rgba(0,0,0,0.08);
        border-color: #CBD5E1;
    }

    .hub-card-title {
        font-size: 25px;
        font-weight: 750;
        color: #111827;
        margin-bottom: 12px;
    }

    .hub-card-text {
        font-size: 15.5px;
        color: #4B5563;
        line-height: 1.55;
        margin-bottom: 18px;
    }

    .tag {
        display: inline-block;
        padding: 7px 11px;
        margin: 4px 4px 0 0;
        border-radius: 999px;
        background: #EEF2FF;
        color: #3730A3;
        font-size: 13px;
        font-weight: 650;
    }

    .note-box {
        margin-top: 28px;
        padding: 17px 20px;
        border-radius: 14px;
        background: #F9FAFB;
        border: 1px solid #E5E7EB;
        color: #374151;
        font-size: 15.5px;
        line-height: 1.5;
    }

    </style>
    """,
    unsafe_allow_html=True
)

# Conteúdo da página inicial
st.markdown(
    """
    <div class="main-title">SCM Analytics Hub</div>
    <div class="subtitle">
        Portal de dashboards para Supply Chain: preço, fornecedores, volumes, custos ponderados e análises operacionais.
    </div>
    """,
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)

with col1:
    st.markdown(
        """
        <div class="hub-card">
            <div class="hub-card-title">🔎 Cockpit Papel</div>
            <div class="hub-card-text">
                Análise de preços, fornecedores, P.Value, TCO, deduplicação e visão limpa de registros comerciais.
            </div>
            <span class="tag">Preço</span>
            <span class="tag">P.Value</span>
            <span class="tag">Fornecedores</span>
            <span class="tag">Auditoria</span>
        </div>
        """,
        unsafe_allow_html=True
    )

with col2:
    st.markdown(
        """
        <div class="hub-card">
            <div class="hub-card-title">📈 Paper Base</div>
            <div class="hub-card-text">
                Volumes mensais, ranking por fornecedor, YTD e custo médio ponderado em EUR/kg.
            </div>
            <span class="tag">Volume mensal</span>
            <span class="tag">EUR/kg</span>
            <span class="tag">YTD</span>
            <span class="tag">Supplier ranking</span>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown(
    """
    <div class="note-box">
        Use o menu lateral para abrir os dashboards.
        Esta página é apenas a tela inicial do SCM Analytics Hub.
    </div>
    """,
    unsafe_allow_html=True
)
