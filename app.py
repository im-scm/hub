
import os
from pathlib import Path
import streamlit as st

st.set_page_config(page_title="SCM Analytics Hub", layout="wide", initial_sidebar_state="expanded")

LOGO_PATH = "assets/impress_logo.png"
PAGE_COCKPIT = "pages/01_Cockpit_Papel.py"
PAGE_PAPERBASE = "pages/02_Paper_Base.py"

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


def page_exists(page_path: str) -> bool:
    return Path(page_path).is_file()


def open_page(page_path: str):
    if not page_exists(page_path):
        st.error("Página não encontrada.")
        st.info(f"O arquivo precisa existir exatamente neste caminho: {page_path}")
        st.code("\n".join(sorted(str(p) for p in Path('.').glob('**/*.py'))))
        st.stop()
    st.switch_page(page_path)


def nav_button(label: str, page_path: str, icon: str):
    exists = page_exists(page_path)
    button_label = f"{icon} {label}" if exists else f"⚠️ {label} não encontrado"
    if st.button(button_label, width="stretch", disabled=not exists):
        open_page(page_path)


with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=145)

    st.markdown("### SCM Analytics Hub")
    st.caption("Navegação")
    st.divider()

    st.button("🏠 Início", width="stretch", disabled=True)
    nav_button("Cockpit Papel", PAGE_COCKPIT, "📊")
    nav_button("Paper Base", PAGE_PAPERBASE, "📦")

    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("Arquivos esperados na raiz: Cockpit_Papel.xlsm e app_paperbase.xlsx")

st.markdown("<div class='hub-title'>SCM Analytics Hub</div>", unsafe_allow_html=True)
st.markdown("<div class='hub-subtitle'>Portal de dashboards para Supply Chain: preço, fornecedores, volumes, custos ponderados e análises operacionais.</div>", unsafe_allow_html=True)

missing_pages = [p for p in [PAGE_COCKPIT, PAGE_PAPERBASE] if not page_exists(p)]
if missing_pages:
    st.warning("Algumas páginas não foram encontradas. Confira a estrutura do repositório no GitHub.")
    st.code("""/
├── app.py
├── requirements.txt
├── Cockpit_Papel.xlsm
├── app_paperbase.xlsx
├── assets/
│   └── impress_logo.png
└── pages/
    ├── 01_Cockpit_Papel.py
    └── 02_Paper_Base.py
""")
    st.markdown("**Páginas faltando:**")
    for p in missing_pages:
        st.markdown(f"- `{p}`")

col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    <div class='hub-card'>
      <div class='hub-card-title'>Cockpit Papel</div>
      <div class='hub-card-text'>Análise de preços, fornecedores, P.Value, TCO, deduplicação e visão limpa de registros comerciais.</div>
      <span class='hub-tag'>Preço</span><span class='hub-tag'>P.Value</span><span class='hub-tag'>Fornecedores</span><span class='hub-tag'>Auditoria</span>
    </div>
    """, unsafe_allow_html=True)
    if page_exists(PAGE_COCKPIT):
        if st.button("📊 Abrir Cockpit Papel", width="stretch"):
            open_page(PAGE_COCKPIT)
    else:
        st.error(f"Arquivo ausente: {PAGE_COCKPIT}")

with col2:
    st.markdown("""
    <div class='hub-card'>
      <div class='hub-card-title'>Paper Base</div>
      <div class='hub-card-text'>Volumes mensais, ranking por fornecedor, YTD e custo médio ponderado em EUR/kg.</div>
      <span class='hub-tag'>Volume mensal</span><span class='hub-tag'>EUR/kg</span><span class='hub-tag'>YTD</span><span class='hub-tag'>Supplier ranking</span>
    </div>
    """, unsafe_allow_html=True)
    if page_exists(PAGE_PAPERBASE):
        if st.button("📦 Abrir Paper Base", width="stretch"):
            open_page(PAGE_PAPERBASE)
    else:
        st.error(f"Arquivo ausente: {PAGE_PAPERBASE}")

with st.expander("Diagnóstico técnico: arquivos .py encontrados"):
    py_files = sorted(str(p) for p in Path('.').glob('**/*.py'))
    if py_files:
        st.code("\n".join(py_files))
    else:
        st.code("Nenhum arquivo .py encontrado além do app principal.")

st.markdown(
    "<div class='hub-note'>Esta versão mostra aviso claro quando a pasta pages ou os arquivos das páginas não existem no caminho esperado.</div>",
    unsafe_allow_html=True,
)
