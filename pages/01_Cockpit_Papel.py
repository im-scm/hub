
import os
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from html import escape

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook

EXCEL_FILE = "Cockpit_Papel.xlsm"
SOURCE_SHEET = "Preços e Condições"
PREMISSAS_SHEET = "Premissas"
APP_TITLE = "Cockpit Papel"
APP_VERSION = "V2.3 | Multipage"
TABLE_HEIGHT_PX = 560

st.markdown("""
<style>
.block-container {padding-top: 1.1rem; padding-bottom: 1.2rem;}
.cp-title {font-size:1.75rem; font-weight:700; color:#111827; margin:0 0 .2rem 0;}
.cp-kpi {border:1px solid #E5E7EB; border-radius:14px; padding:14px 16px; background:#FFF; box-shadow:0 1px 2px rgba(0,0,0,.04); min-height:82px;}
.cp-kpi-label {font-size:.78rem; color:#6B7280; margin-bottom:6px;}
.cp-kpi-value {font-size:1.14rem; font-weight:700; color:#111827; line-height:1.2;}
.cp-table-shell {border:1px solid #E5E7EB; border-radius:14px; background:#FFF; overflow:hidden; box-shadow:0 1px 2px rgba(0,0,0,.03);}
.cp-table-wrap {overflow-y:auto; overflow-x:auto;}
.cp-table {width:100%; border-collapse:separate; border-spacing:0; table-layout:auto; font-size:.84rem; color:#111827;}
.cp-table thead th {position:sticky; top:0; z-index:2; background:#F9FAFB; color:#6B7280; font-weight:600; border-bottom:1px solid #E5E7EB; border-right:1px solid #F0F2F5; padding:8px 10px; white-space:nowrap;}
.cp-table tbody td {padding:7px 10px; border-bottom:1px solid #F3F4F6; border-right:1px solid #F8FAFC; white-space:nowrap; background:#FFF;}
.cp-table tbody tr:hover td {background:#FAFBFF;}
.cp-left{text-align:left}.cp-center{text-align:center}.cp-right{text-align:right;font-variant-numeric:tabular-nums}.cp-note{color:#6B7280;font-size:.78rem;margin-top:.35rem;}
</style>
""", unsafe_allow_html=True)


def normalize_text(value):
    if value is None or pd.isna(value): return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", text)

def parse_number(value):
    if pd.isna(value): return None
    if isinstance(value, (int, float)): return float(value)
    s = str(value).strip().replace("R$", "").replace("€", "").replace(" ", "").replace("%", "")
    s = re.sub(r"[^0-9,.-]", "", s)
    if s in {"", "-", ".", ","}: return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    elif "," in s: s = s.replace(".", "").replace(",", ".")
    try: return float(s)
    except Exception: return None

def format_br_number(value, decimals=2):
    if pd.isna(value): return ""
    try: n = float(value)
    except Exception: return str(value)
    return f"{n:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_no_decimal(value):
    if pd.isna(value): return ""
    try: return f"{float(value):,.0f}".replace(",", ".")
    except Exception: return str(value)

def find_column(columns, aliases):
    norm = {c: normalize_text(c) for c in columns}
    for alias in aliases:
        a = normalize_text(alias)
        for c, n in norm.items():
            if n == a: return c
        for c, n in norm.items():
            if a in n: return c
    return None

def detect_header_row(raw):
    targets = ["impress type", "supplier", "current price"]
    best, score_best = None, -1
    for i in range(min(len(raw), 50)):
        text = " | ".join(normalize_text(v) for v in raw.iloc[i].tolist())
        score = sum(t in text for t in targets)
        if score > score_best: best, score_best = i, score
    return best if score_best >= 2 else None

def kpi_card(label, value):
    st.markdown(f"<div class='cp-kpi'><div class='cp-kpi-label'>{escape(str(label))}</div><div class='cp-kpi-value'>{escape(str(value))}</div></div>", unsafe_allow_html=True)

def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Cockpit Filtrado")
    output.seek(0)
    return output.getvalue()

def build_html_table(df_html, height_px=560):
    if df_html.empty: return "<div class='cp-note'>Sem registros para exibir.</div>"
    right = {"Current Price", "P.Value (R$/KG)", "P.Value (R$/M2)", "Payment Terms", "Lot (ton)", "g/m2"}
    left = {"Supplier"}
    parts = ["<div class='cp-table-shell'>", f"<div class='cp-table-wrap' style='max-height:{height_px}px;'>", "<table class='cp-table'><thead><tr>"]
    for col in df_html.columns:
        cls = "cp-right" if col in right else ("cp-left" if col in left else "cp-center")
        parts.append(f"<th class='{cls}'>{escape(str(col))}</th>")
    parts.append("</tr></thead><tbody>")
    for _, row in df_html.iterrows():
        parts.append("<tr>")
        for col in df_html.columns:
            cls = "cp-right" if col in right else ("cp-left" if col in left else "cp-center")
            value = "" if pd.isna(row[col]) else str(row[col])
            parts.append(f"<td class='{cls}'>{escape(value)}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div></div>")
    return "".join(parts)

def load_premissas():
    try:
        p = pd.read_excel(EXCEL_FILE, sheet_name=PREMISSAS_SHEET, header=None, engine="openpyxl")
        def cell(r,c):
            try:
                v=p.iat[r,c]
                return None if pd.isna(v) else v
            except Exception: return None
        return {"Frete CN":cell(5,2),"Frete EU":cell(4,2),"USD/BRL":cell(23,2),"EUR/BRL":cell(24,2),"CNY/BRL":cell(25,2)}
    except Exception:
        return {"Frete CN":None,"Frete EU":None,"USD/BRL":None,"EUR/BRL":None,"CNY/BRL":None}

def get_excel_last_update(file_path):
    try:
        wb = load_workbook(file_path, read_only=True, keep_vba=True)
        m = wb.properties.modified
        return m.strftime("%d/%m/%Y") if m else "N/A"
    except Exception:
        try: return datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%d/%m/%Y")
        except Exception: return "N/A"

def deduplicate(df):
    d = df.copy()
    date_col = "Última Atualização de Preço"
    if date_col in d.columns:
        d[date_col] = pd.to_datetime(d[date_col], errors="coerce", dayfirst=True)
        d = d.sort_values(date_col, ascending=False, na_position="last", kind="stable")
    keys = [c for c in ["Impress Type","g/m2","Supplier","Current Price","Currency","P.Value (R$/KG)","P.Value (R$/M2)","Payment Terms","Lot (ton)"] if c in d.columns]
    for c in keys:
        if pd.api.types.is_object_dtype(d[c]): d[c] = d[c].apply(normalize_text)
    if keys: d = d.drop_duplicates(keys, keep="first")
    return d.reset_index(drop=True)

def build_canonical(df):
    aliases = {
        "Material Number":["Material Number","Material"],"Impress Type":["Impress Type","Print Type"],"Width (mm)":["Width (mm)","Width"],"g/m2":["g/m2","g/m²","Gramatura"],"Supplier":["Supplier","Fornecedor"],"Currency":["Currency","Moeda"],"Current Price":["Current Price","Preço Atual","Preco Atual"],"Lot (ton)":["Lot (ton)","Lot","Lote"],"TCO (R$/KG)":["TCO (R$/KG)","TCO"],"TCO (R$/M2)":["TCO (R$/M2)"],"Payment Terms":["Payment Terms","Prazo Pagamento"],"P.Value (R$/KG)":["P.Value (R$/KG)","P.Value"],"P.Value (R$/M2)":["P.Value (R$/M2)"],"Última Atualização de Preço":["Última Atualização de Preço","Ultima Atualizacao de Preco","Last Update"]
    }
    rename = {}
    for k,a in aliases.items():
        orig = find_column(df.columns, a)
        if orig: rename[orig]=k
    d = df.rename(columns=rename).copy()
    cols = [c for c in aliases.keys() if c in d.columns]
    d = d[cols].copy()
    nums = ["Width (mm)","g/m2","Current Price","Lot (ton)","TCO (R$/KG)","TCO (R$/M2)","P.Value (R$/KG)","P.Value (R$/M2)","Payment Terms"]
    for c in nums:
        if c in d.columns: d[c]=d[c].apply(parse_number)
    if "Última Atualização de Preço" in d.columns: d["Última Atualização de Preço"] = pd.to_datetime(d["Última Atualização de Preço"], errors="coerce", dayfirst=True)
    if "TCO (R$/KG)" not in d.columns and "Current Price" in d.columns: d["TCO (R$/KG)"]=d["Current Price"]
    if "P.Value (R$/KG)" not in d.columns and "TCO (R$/KG)" in d.columns: d["P.Value (R$/KG)"]=d["TCO (R$/KG)"]
    if "TCO (R$/M2)" not in d.columns and {"TCO (R$/KG)","g/m2"}.issubset(d.columns): d["TCO (R$/M2)"]=d["TCO (R$/KG)"]*(d["g/m2"]/1000)
    if "P.Value (R$/M2)" not in d.columns and {"P.Value (R$/KG)","g/m2"}.issubset(d.columns): d["P.Value (R$/M2)"]=d["P.Value (R$/KG)"]*(d["g/m2"]/1000)
    for c in ["Impress Type","Supplier"]:
        if c in d.columns: d = d.dropna(subset=[c])
    if "Current Price" in d.columns: d=d[d["Current Price"].notna()]
    d = d.drop(columns=[c for c in ["Material Number","Width (mm)"] if c in d.columns])
    return deduplicate(d)

@st.cache_data(show_spinner=False)
def load_data_with_audit():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=None, engine="openpyxl")
    h = detect_header_row(raw)
    if h is None: raise ValueError("Não foi possível identificar o cabeçalho da aba 'Preços e Condições'.")
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=h, engine="openpyxl")
    df_raw.columns = [re.sub(r"\s+", " ", str(c).strip().replace("\n", " ")) for c in df_raw.columns]
    df_raw = df_raw.dropna(how="all")
    final = build_canonical(df_raw)
    return final, {"original":len(df_raw),"final":len(final),"removed":max(len(df_raw)-len(final),0)}

def create_filter(data, col, label, numeric=False):
    if col not in data.columns: return data
    vals = data[col].dropna().unique().tolist()
    if numeric: vals = sorted(vals)
    else: vals = sorted([str(v) for v in vals])
    selected = st.sidebar.multiselect(label, vals, default=[])
    if not selected: return data
    if numeric: return data[data[col].isin(selected)]
    return data[data[col].astype(str).isin(selected)]

def export_table(data):
    cols = ["Impress Type","g/m2","Supplier","Current Price","Currency","P.Value (R$/KG)","P.Value (R$/M2)","Última Atualização de Preço","Payment Terms","Lot (ton)"]
    d = data[[c for c in cols if c in data.columns]].copy().rename(columns={"Última Atualização de Preço":"Último Preço"})
    if "P.Value (R$/M2)" in d.columns:
        sort = pd.to_numeric(data["P.Value (R$/M2)"], errors="coerce")
        d = d.assign(_sort=sort.values).sort_values("_sort", na_position="last", kind="stable").drop(columns="_sort")
    for c in ["Current Price","P.Value (R$/KG)","P.Value (R$/M2)"]:
        if c in d.columns: d[c]=d[c].apply(lambda x: format_br_number(x,2))
    for c in ["g/m2","Payment Terms","Lot (ton)"]:
        if c in d.columns: d[c]=d[c].apply(format_no_decimal)
    if "Último Preço" in d.columns: d["Último Preço"]=pd.to_datetime(d["Último Preço"], errors="coerce").dt.strftime("%d/%m/%Y")
    return d

try:
    df, audit = load_data_with_audit()
except Exception as e:
    st.error("❌ Falha ao carregar Cockpit_Papel.xlsm.")
    st.code(str(e)); st.stop()

prem = load_premissas()
st.markdown(f"<div class='cp-title'>{APP_TITLE}</div>", unsafe_allow_html=True)
st.caption("Análise de preços, fornecedores e P.Value com tabela ordenada por menor P.Value (R$/M2).")

with st.sidebar:
    st.markdown("### Filtros")
    st.caption(APP_VERSION)
    st.divider()
filtered = df.copy()
filtered = create_filter(filtered,"Impress Type","Impress Type")
filtered = create_filter(filtered,"g/m2","g/m2", numeric=True)
filtered = create_filter(filtered,"Supplier","Supplier")
filtered = create_filter(filtered,"Currency","Currency")
filtered = create_filter(filtered,"Lot (ton)","Lot (ton)", numeric=True)
st.sidebar.caption(f"Última atualização arquivo: {get_excel_last_update(EXCEL_FILE)}")

if filtered.shape[0] == df.shape[0]:
    st.info("Selecione pelo menos um filtro para visualizar os dados."); st.stop()

cols = st.columns(5)
for col, label in zip(cols, ["Frete CN","Frete EU","USD/BRL","EUR/BRL","CNY/BRL"]):
    with col: kpi_card(label, format_br_number(prem[label],2) if prem[label] is not None else "N/A")

a1,a2,a3 = st.columns(3)
with a1: kpi_card("Linhas originais", format_no_decimal(audit["original"]))
with a2: kpi_card("Linhas removidas", format_no_decimal(audit["removed"]))
with a3: kpi_card("Linhas finais", format_no_decimal(audit["final"]))

table = export_table(filtered)
st.markdown("### Tabela principal")
st.markdown(build_html_table(table, TABLE_HEIGHT_PX), unsafe_allow_html=True)
st.markdown("<div class='cp-note'>Tabela ordenada do menor para o maior em P.Value (R$/M2).</div>", unsafe_allow_html=True)

b1,b2,_=st.columns([1.2,1.2,6])
with b1: st.download_button("Exportar CSV", table.to_csv(index=False,sep=';',encoding='utf-8-sig').encode('utf-8-sig'), "cockpit_filtrado.csv", "text/csv", width="stretch")
with b2: st.download_button("Exportar Excel", to_excel_bytes(table), "cockpit_filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

st.markdown("### TCO médio por Supplier")
if {"Supplier","TCO (R$/KG)"}.issubset(filtered.columns):
    chart = filtered.dropna(subset=["Supplier","TCO (R$/KG)"]).copy()
    chart["TCO (R$/KG)"]=pd.to_numeric(chart["TCO (R$/KG)"], errors="coerce")
    chart = chart.groupby("Supplier", as_index=False)["TCO (R$/KG)"].mean().sort_values("TCO (R$/KG)")
    chart["label"] = chart["TCO (R$/KG)"].apply(lambda x: format_br_number(x,2))
    fig = px.bar(chart, x="Supplier", y="TCO (R$/KG)", color="TCO (R$/KG)", color_continuous_scale="Blues", text="label")
    fig.update_traces(textposition="outside")
    fig.update_layout(showlegend=False, height=420, margin=dict(t=10,r=20,l=20,b=20), plot_bgcolor="white", paper_bgcolor="white")
    st.plotly_chart(fig, width="stretch")
