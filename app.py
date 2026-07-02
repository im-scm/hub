
import os
import runpy
from pathlib import Path
import streamlit as st

st.set_page_config(page_title="SCM Analytics Hub", layout="wide", initial_sidebar_state="expanded")

LOGO_PATH = "assets/impress_logo.png"
PAGE_FILES = {
    "Cockpit Papel": "pages/01_Cockpit_Papel.py",
    "Paper Base": "pages/02_Paper_Base.py",
}

st.markdown("""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 1.2rem;}
.hub-title {font-size: 2.0rem; font-weight: 750; color: #111827; margin-bottom: .2rem;}
.hub-subtitle {font-size: .98rem; color: #6B7280; margin-bottom: 1.2rem;}
.hub-card {border: 1px solid #E5E7EB; border-radius: 18px; padding: 20px 22px; background: #FFFFFF; box-shadow: 0 1px 3px rgba(0,0,0,.05); min-height: 170px;}
.hub-card-title {font-size: 1.15rem; font-weight: 700; color: #111827; margin-bottom: .45rem;}
.hub-card-text {font-size: .90rem; color: #4B5563; line-height: 1.45; margin-bottom: .8rem;}
.hub-tag {display: inline-block; font-size: .75rem; color: #374151; background: #F3F4F6; border: 1px solid #E5E7EB; border-radius: 999px; padding: 4px 9px; margin-right: 5px; margin-top: 4px;}
.hub-note {font-size: .82rem; color: #6B7280; margin-top: 1.2rem;}
</style>
""", unsafe_allow_html=True)


def render_home():
    st.markdown("<div class='hub-title'>SCM Analytics Hub</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='hub-subtitle'>Portal de dashboards para Supply Chain: preço, fornecedores, volumes, custos ponderados e análises operacionais.</div>",
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class='hub-card'>
          <div class='hub-card-title'>Cockpit Papel</div>
          <div class='hub-card-text'>Análise de preços, fornecedores, P.Value, TCO, deduplicação e visão limpa de registros comerciais.</div>
          <span class='hub-tag'>Preço</span><span class='hub-tag'>P.Value</span><span class='hub-tag'>Fornecedores</span><span class='hub-tag'>Auditoria</span>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='hub-card'>
          <div class='hub-card-title'>Paper Base</div>
          <div class='hub-card-text'>Volumes mensais, ranking por fornecedor, YTD e custo médio ponderado em EUR/kg.</div>
          <span class='hub-tag'>Volume mensal</span><span class='hub-tag'>EUR/kg</span><span class='hub-tag'>YTD</span><span class='hub-tag'>Supplier ranking</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown(
        "<div class='hub-note'>Use o seletor na sidebar para abrir cada dashboard. Esta versão não usa st.page_link nem st.switch_page.</div>",
        unsafe_allow_html=True,
    )


def run_dashboard(page_file: str):
    path = Path(page_file)
    if not path.is_file():
        st.error("Arquivo da página não encontrado.")
        st.info(f"Caminho esperado: {page_file}")
        with st.expander("Arquivos .py encontrados no repositório"):
            found = sorted(str(p) for p in Path('.').glob('**/*.py'))
            st.code("\n".join(found) if found else "Nenhum arquivo .py encontrado.")
        st.stop()
    runpy.run_path(str(path), run_name="__main__")


with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=145)

    st.markdown("### SCM Analytics Hub")
    st.caption("Navegação estável sem switch_page")
    st.divider()

    selected_page = st.radio(
        "Selecione a visão",
        options=["Início", "Cockpit Papel", "Paper Base"],
        index=0,
        label_visibility="collapsed",
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("Arquivos Excel na raiz:")
    st.caption("Cockpit_Papel.xlsm")
    st.caption("app_paperbase.xlsx")

if selected_page == "Início":
    render_home()
else:
    run_dashboard(PAGE_FILES[selected_page])
