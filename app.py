
import os
import streamlit as st

st.set_page_config(page_title="SCM Analytics Hub", layout="wide", initial_sidebar_state="expanded")

LOGO_PATH = "assets/impress_logo.png"

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 1.2rem;}
.hub-title {font-size: 2.0rem; font-weight: 750; color: #111827; margin-bottom: .2rem;}
.hub-subtitle {font-size: .98rem; color: #6B7280; margin-bottom: 1.2rem;}
.hub-card {border: 1px solid #E5E7EB; border-radius: 18px; padding: 20px 22px; background: #FFFFFF; box-shadow: 0 1px 3px rgba(0,0,0,.05); min-height: 170px;}
.hub-card-title {font-size: 1.15rem; font-weight: 700; color: #111827; margin-bottom: .45rem;}
.hub-card-text {font-size: .90rem; color: #4B5563; line-height: 1.45; margin-bottom: .8rem;}
.hub-tag {display: inline-block; font-size: .75rem; color: #374151; background: #F3F4F6; border: 1px solid #E5E7EB; border-radius: 999px; padding: 4px 9px; margin-right: 5px; margin-top: 4px;}
.hub-note {font-size: .82rem; color: #6B7280; margin-top: 1.2rem;}
</style>
""", unsafe_allow_html=True)

def go_to_page(page_path):
    try:
        st.switch_page(page_path)
    except Exception as e:
        st.error("Não consegui abrir a página automaticamente.")
        st.code(str(e))
        st.info("Confira se o arquivo existe exatamente neste caminho: " + page_path)

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=145)

    st.markdown("### SCM Analytics Hub")
    st.caption("Navegação")
    st.divider()

    if st.button("🏠 Início", width="stretch"):
        st.rerun()
    if st.button("📊 Cockpit Papel", width="stretch"):
        go_to_page("pages/01_Cockpit_Papel.py")
    if st.button("📦 Paper Base", width="stretch"):
        go_to_page("pages/02_Paper_Base.py")

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("Arquivos esperados na raiz: Cockpit_Papel.xlsm e app_paperbase.xlsx")

st.markdown("<div class='hub-title'>SCM Analytics Hub</div>", unsafe_allow_html=True)
st.markdown("<div class='hub-subtitle'>Portal de dashboards para Supply Chain: preço, fornecedores, volumes, custos ponderados e análises operacionais.</div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    <div class='hub-card'>
      <div class='hub-card-title'>Cockpit Papel</div>
      <div class='hub-card-text'>Análise de preços, fornecedores, P.Value, TCO, deduplicação e visão limpa de registros comerciais.</div>
      <span class='hub-tag'>Preço</span><span class='hub-tag'>P.Value</span><span class='hub-tag'>Fornecedores</span><span class='hub-tag'>Auditoria</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📊 Abrir Cockpit Papel", width="stretch"):
        go_to_page("pages/01_Cockpit_Papel.py")

with col2:
    st.markdown("""
    <div class='hub-card'>
      <div class='hub-card-title'>Paper Base</div>
      <div class='hub-card-text'>Volumes mensais, ranking por fornecedor, YTD e custo médio ponderado em EUR/kg.</div>
      <span class='hub-tag'>Volume mensal</span><span class='hub-tag'>EUR/kg</span><span class='hub-tag'>YTD</span><span class='hub-tag'>Supplier ranking</span>
    </div>
    """, unsafe_allow_html=True)
    if st.button("📦 Abrir Paper Base", width="stretch"):
        go_to_page("pages/02_Paper_Base.py")

st.markdown(
    "<div class='hub-note'>Esta versão evita st.page_link, que estava gerando KeyError: url_pathname no Streamlit Cloud. A navegação usa st.switch_page acionado por botões.</div>",
    unsafe_allow_html=True,
)
