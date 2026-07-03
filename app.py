from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="SCM Analytics Hub",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>

/* Home */
[data-testid="stSidebarNav"] ul li:nth-child(1) a {
    font-size: 0;
}

[data-testid="stSidebarNav"] ul li:nth-child(1) a::after {
    content: "🏠 Inicio";
    font-size: 14px;
    font-weight: 600;
}

/* Cockpit Papel */
[data-testid="stSidebarNav"] ul li:nth-child(2) a {
    font-size: 0;
}

[data-testid="stSidebarNav"] ul li:nth-child(2) a::after {
    content: "🔎 Cockpit Papel";
    font-size: 14px;
    font-weight: 600;
}

/* Paper Base */
[data-testid="stSidebarNav"] ul li:nth-child(3) a {
    font-size: 0;
}

[data-testid="stSidebarNav"] ul li:nth-child(3) a::after {
    content: "📈 Paper Base";
    font-size: 14px;
    font-weight: 600;
}

</style>
""", unsafe_allow_html=True)

LOGO_PATH = "assets/impress_logo.png"

# Logo na sidebar, se existir
if Path(LOGO_PATH).is_file():
    st.logo(LOGO_PATH, size="large")

# CSS leve apenas para a home
st.markdown(
    """
    <style>
        .main-title {
            font-size: 42px;
            font-weight: 800;
            color: #111827;
            margin-bottom: 6px;
        }

        .subtitle {
            font-size: 18px;
            color: #4B5563;
            margin-bottom: 28px;
        }

        .hub-card {
            border: 1px solid #E5E7EB;
            border-radius: 16px;
            padding: 24px;
            background: #FFFFFF;
            box-shadow: 0 2px 10px rgba(0,0,0,0.04);
            min-height: 210px;
        }

        .hub-card-title {
            font-size: 24px;
            font-weight: 700;
            color: #111827;
            margin-bottom: 10px;
        }

        .hub-card-text {
            font-size: 15px;
            color: #4B5563;
            line-height: 1.5;
            margin-bottom: 16px;
        }

        .tag {
            display: inline-block;
            padding: 6px 10px;
            margin: 4px 4px 0 0;
            border-radius: 999px;
            background: #EEF2FF;
            color: #3730A3;
            font-size: 13px;
            font-weight: 600;
        }

        .note-box {
            margin-top: 26px;
            padding: 16px 18px;
            border-radius: 12px;
            background: #F9FAFB;
            border: 1px solid #E5E7EB;
            color: #374151;
            font-size: 15px;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# Home principal
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
            <div class="hub-card-title">Cockpit Papel</div>
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
            <div class="hub-card-title">Paper Base</div>
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
