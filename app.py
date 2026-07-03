from pathlib import Path
import streamlit as st

st.set_page_config(
    page_title="SCM Analytics Hub",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>

/* Esconde o texto original dos links da navegação */
[data-testid="stSidebarNav"] ul li a span {
    display: none !important;
}

/* Mantém os links com boa altura e alinhamento */
[data-testid="stSidebarNav"] ul li a {
    display: flex !important;
    align-items: center !important;
    min-height: 38px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    text-decoration: none !important;
}

/* Página inicial: app.py */
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

</style>
""", unsafe_allow_html=True)

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
