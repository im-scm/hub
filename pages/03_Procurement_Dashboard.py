import os
import re
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="Procurement Dashboard", layout="wide")

<style>

/* Esconde o texto original da navegação lateral */
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

/* Inicio */
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

/* Procurement Dashboard */
    [data-testid="stSidebarNav"] ul li:nth-child(4) a::after {
        content: "🧭 Procurement Dashboard";
        font-size: 14px;
        font-weight: 600;
}

</style>    

APP_TITLE = "Procurement Dashboard"
APP_VERSION = "v2.0 | Procurement Analytics"
LOGO_PATH = "assets/impress_logo.png"
DATA_CANDIDATES = [
    "data/ksoft.xlsx",
    "ksoft.xlsx",
    "database/ksoft.xlsx",
]

if os.path.exists(LOGO_PATH):
    st.logo(LOGO_PATH, size="large")

# -----------------------------------------------------------------------------
# CSS leve, alinhado ao hub atual
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.0rem; padding-bottom: 1.2rem;}
    .scm-title {font-size: 2.05rem; font-weight: 760; color:#111827; margin-bottom:.15rem;}
    .scm-subtitle {font-size:.98rem; color:#6B7280; margin-bottom:1.15rem;}
    .kpi-card {
        border:1px solid #E5E7EB; border-radius:14px; padding:14px 16px;
        background:linear-gradient(180deg,#FFFFFF 0%,#F9FAFB 100%);
        box-shadow:0 1px 2px rgba(15,23,42,.06); min-height:92px;
    }
    .kpi-label {font-size:.77rem; text-transform:uppercase; color:#6B7280; letter-spacing:.04em; font-weight:700;}
    .kpi-value {font-size:1.45rem; color:#111827; font-weight:780; margin-top:.35rem;}
    .kpi-note {font-size:.77rem; color:#6B7280; margin-top:.25rem;}
    .section-title {font-size:1.06rem; font-weight:760; color:#111827; margin: .35rem 0 .4rem;}
    .small-muted {font-size:.82rem; color:#6B7280;}
    div[data-testid="stMetricValue"] {font-size: 1.28rem;}
    .dataframe th {font-size: 0.8rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Utilitários
# -----------------------------------------------------------------------------
def fmt_brl(value, decimals=0):
    if pd.isna(value):
        return "-"
    try:
        n = float(value)
    except Exception:
        return str(value)
    s = f"R$ {n:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(value, decimals=0):
    if pd.isna(value):
        return "-"
    try:
        n = float(value)
    except Exception:
        return str(value)
    s = f"{n:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value, decimals=1):
    if pd.isna(value):
        return "-"
    try:
        return f"{float(value)*100:.{decimals}f}%".replace(".", ",")
    except Exception:
        return str(value)


def kpi_card(label, value, note=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def clean_columns(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def to_numeric(series):
    return pd.to_numeric(series, errors="coerce")


def excel_date_to_datetime(value):
    """Converte datas vindas como serial Excel, texto ou timestamp."""
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.to_datetime(value, errors="coerce")
    # Excel serial típico: 46000 etc.
    if isinstance(value, (int, float, np.integer, np.floating)):
        if 20000 <= float(value) <= 90000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
    # string numérica
    text = str(value).strip()
    try:
        n = float(text)
        if 20000 <= n <= 90000:
            return pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
    except Exception:
        pass
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def find_data_file():
    for p in DATA_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def read_excel_source(uploaded_file=None):
    source = uploaded_file if uploaded_file is not None else find_data_file()
    if source is None:
        return None, None, None, None
    t001 = pd.read_excel(source, sheet_name="T001", engine="openpyxl")
    supplier_group = pd.read_excel(source, sheet_name="Supplier Group", engine="openpyxl")
    payment_terms = pd.read_excel(source, sheet_name="Payment Terms", engine="openpyxl")
    source_name = getattr(uploaded_file, "name", source)
    return t001, supplier_group, payment_terms, source_name


@st.cache_data(show_spinner=False)
def load_from_path(path):
    t001 = pd.read_excel(path, sheet_name="T001", engine="openpyxl")
    supplier_group = pd.read_excel(path, sheet_name="Supplier Group", engine="openpyxl")
    payment_terms = pd.read_excel(path, sheet_name="Payment Terms", engine="openpyxl")
    return prepare_data(t001, supplier_group, payment_terms)


def load_from_upload(uploaded):
    t001 = pd.read_excel(uploaded, sheet_name="T001", engine="openpyxl")
    supplier_group = pd.read_excel(uploaded, sheet_name="Supplier Group", engine="openpyxl")
    payment_terms = pd.read_excel(uploaded, sheet_name="Payment Terms", engine="openpyxl")
    return prepare_data(t001, supplier_group, payment_terms)


def prepare_data(t001, supplier_group, payment_terms):
    df = clean_columns(t001)
    sg = clean_columns(supplier_group)
    pt = clean_columns(payment_terms)

    # Gerencia colunas obrigatórias de forma defensiva
    numeric_cols = [
        "Order Qua.", "Pos. Value", "Pos. Open Qua", "Price H. C.", "Price",
        "Conv. Rate", "Pos. Del. Qua.", "Pos. Del. Value", "Pos. Inv. Qua.",
        "Maturity days"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = to_numeric(df[col])

    for col in ["Order Date", "Desired Date", "Conf. Date", "Delivery Date", "Conv. Date"]:
        if col in df.columns:
            df[col] = df[col].apply(excel_date_to_datetime)

    # Supplier Group: T001.Supplier Group -> Supplier Group.Group number
    if {"Supplier Group", "Group number", "Description"}.issubset(set(df.columns) | set(sg.columns)):
        sg2 = sg[["Group number", "Description"]].copy()
        sg2["Group number"] = sg2["Group number"].astype(str).str.strip()
        df["_supplier_group_key"] = df["Supplier Group"].astype(str).str.strip()
        df = df.merge(sg2, left_on="_supplier_group_key", right_on="Group number", how="left")
        df = df.rename(columns={"Description": "Supplier Group Description"})
    else:
        df["Supplier Group Description"] = "N/A"

    # Payment Terms: T001.Paymentterms -> Payment Terms.Payment Term
    if {"Paymentterms"}.issubset(df.columns) and {"Payment Term", "Payment Term Description", "Maturity days"}.issubset(pt.columns):
        pt2 = pt[["Payment Term", "Payment Term Description", "Maturity days"]].copy()
        pt2["Payment Term"] = pt2["Payment Term"].astype(str).str.strip()
        pt2["Maturity days"] = to_numeric(pt2["Maturity days"])
        df["_payment_key"] = df["Paymentterms"].astype(str).str.strip()
        df = df.merge(pt2, left_on="_payment_key", right_on="Payment Term", how="left")
    else:
        df["Payment Term Description"] = "N/A"
        df["Maturity days"] = np.nan

    # Campos derivados
    today = pd.Timestamp.today().normalize()
    df["Spend BRL"] = to_numeric(df.get("Pos. Value", 0)).fillna(0)
    df["Open Qty"] = to_numeric(df.get("Pos. Open Qua", 0)).fillna(0)
    df["Delivered Qty"] = to_numeric(df.get("Pos. Del. Qua.", 0)).fillna(0)
    df["Order Qty"] = to_numeric(df.get("Order Qua.", 0)).fillna(0)
    df["Open Value Est."] = np.where(df["Order Qty"] > 0, df["Spend BRL"] * (df["Open Qty"] / df["Order Qty"]), 0)
    df["Open Value Est."] = pd.to_numeric(df["Open Value Est."], errors="coerce").fillna(0)
    df["Fill Rate"] = np.where(df["Order Qty"] > 0, df["Delivered Qty"] / df["Order Qty"], np.nan)
    df["Is Open"] = df["Open Qty"] > 0
    df["Is Overdue"] = df["Is Open"] & df["Conf. Date"].notna() & (df["Conf. Date"] < today)
    df["Is Risk"] = df["Is Open"] & df["Desired Date"].notna() & df["Conf. Date"].notna() & (df["Conf. Date"] > df["Desired Date"])
    df["Days Late"] = np.where(df["Is Overdue"], (today - df["Conf. Date"]).dt.days, 0)
    df["Order Month"] = df["Order Date"].dt.to_period("M").astype(str)
    df["Order Year"] = df["Order Date"].dt.year
    df["Supplier Group Description"] = df["Supplier Group Description"].fillna("Sem grupo")
    df["Supplier Name"] = df["Supplier Name"].fillna("N/A")
    df["Clerk name"] = df["Clerk name"].fillna("N/A")
    df["Currency"] = df["Currency"].fillna("N/A")
    df["Supplier's Country"] = df["Supplier's Country"].fillna("N/A")
    df["Part No."] = df["Part No."].astype(str).fillna("N/A")
    df["Name"] = df["Name"].fillna("N/A")

    return df


def abc_classification(df, group_col, value_col="Spend BRL"):
    if df.empty or group_col not in df.columns:
        return pd.DataFrame(columns=[group_col, value_col, "Share %", "Cum %", "ABC"])
    base = (
        df.groupby(group_col, dropna=False, as_index=False)
        .agg(**{value_col: (value_col, "sum"), "Pedidos": ("Order No.", "nunique"), "Linhas": ("Order No.", "count")})
        .sort_values(value_col, ascending=False)
        .reset_index(drop=True)
    )
    total = base[value_col].sum()
    base["Share %"] = np.where(total > 0, base[value_col] / total, 0)
    base["Cum %"] = base["Share %"].cumsum()
    base["ABC"] = np.select([base["Cum %"] <= 0.80, base["Cum %"] <= 0.95], ["A", "B"], default="C")
    return base


def format_table_currency(df, currency_cols=None, pct_cols=None, num_cols=None):
    out = df.copy()
    for col in currency_cols or []:
        if col in out.columns:
            out[col] = out[col].apply(lambda x: fmt_brl(x, 0))
    for col in pct_cols or []:
        if col in out.columns:
            out[col] = out[col].apply(lambda x: fmt_pct(x, 1))
    for col in num_cols or []:
        if col in out.columns:
            out[col] = out[col].apply(lambda x: fmt_num(x, 0))
    return out


def monthly_spend_chart(df):
    base = df.groupby("Order Month", as_index=False)["Spend BRL"].sum().sort_values("Order Month")
    fig = px.bar(base, x="Order Month", y="Spend BRL", text="Spend BRL", title="Spend mensal")
    fig.update_traces(marker_color="#2563EB", texttemplate="%{text:,.0f}", textposition="outside")
    fig.update_layout(height=410, margin=dict(t=55, r=18, l=18, b=40), plot_bgcolor="white", paper_bgcolor="white", yaxis_title="BRL", xaxis_title="")
    fig.update_yaxes(showgrid=True, gridcolor="#E5E7EB")
    fig.update_xaxes(showgrid=False)
    return fig


def pareto_chart(base, label_col, value_col="Spend BRL", title="Pareto"):
    data = base.head(20).copy()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=data[label_col].astype(str), y=data[value_col], name="Spend", marker_color="#2563EB"), secondary_y=False)
    fig.add_trace(go.Scatter(x=data[label_col].astype(str), y=data["Cum %"] * 100, name="Acumulado %", mode="lines+markers", line=dict(color="#F97316", width=3)), secondary_y=True)
    fig.update_layout(title=title, height=440, margin=dict(t=55, r=50, l=30, b=120), plot_bgcolor="white", paper_bgcolor="white", legend=dict(orientation="h", y=1.08, x=0))
    fig.update_yaxes(title_text="BRL", secondary_y=False, showgrid=True, gridcolor="#E5E7EB")
    fig.update_yaxes(title_text="Acumulado %", secondary_y=True, range=[0, 105])
    fig.update_xaxes(tickangle=-35)
    return fig


def price_changes(df):
    needed = {"Part No.", "Name", "Order Date", "Price H. C."}
    if not needed.issubset(df.columns):
        return pd.DataFrame()
    tmp = df.dropna(subset=["Order Date", "Price H. C."]).copy()
    tmp["Price H. C."] = to_numeric(tmp["Price H. C."])
    tmp = tmp.dropna(subset=["Price H. C."])
    tmp = tmp.sort_values(["Part No.", "Order Date"])
    rows = []
    for part, g in tmp.groupby("Part No."):
        g2 = g.drop_duplicates(subset=["Order Date", "Price H. C."], keep="last").sort_values("Order Date")
        if len(g2) < 2:
            continue
        last = g2.iloc[-1]
        prev = g2.iloc[-2]
        prev_price = prev["Price H. C."]
        last_price = last["Price H. C."]
        if pd.notna(prev_price) and prev_price != 0:
            rows.append({
                "Part No.": part,
                "Name": last.get("Name", ""),
                "Fornecedor": last.get("Supplier Name", ""),
                "Data anterior": prev["Order Date"],
                "Preço anterior": prev_price,
                "Última data": last["Order Date"],
                "Último preço": last_price,
                "Variação %": (last_price / prev_price) - 1,
                "Variação BRL": last_price - prev_price,
            })
    return pd.DataFrame(rows).sort_values("Variação %", ascending=False) if rows else pd.DataFrame()


def export_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Export")
    output.seek(0)
    return output.getvalue()

# -----------------------------------------------------------------------------
# Header e origem dos dados
# -----------------------------------------------------------------------------
st.markdown(f'<div class="scm-title">{APP_TITLE}</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="scm-subtitle">Spend, follow-up, ABC, fornecedores, compradores, PMP e inteligência de preços.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.divider()
    st.markdown("### Base de dados")

uploaded = st.sidebar.file_uploader("Opcional: carregar ksoft.xlsx manualmente", type=["xlsx"])

try:
    if uploaded is not None:
        df = load_from_upload(uploaded)
        source_label = uploaded.name
    else:
        data_path = find_data_file()
        if data_path is None:
            st.warning("Não encontrei o arquivo ksoft.xlsx. Coloque o arquivo em data/ksoft.xlsx ou use o upload na sidebar.")
            st.stop()
        df = load_from_path(data_path)
        source_label = data_path
except Exception as e:
    st.error("Falha ao carregar a base ksoft.xlsx.")
    st.code(str(e))
    st.stop()

with st.sidebar:
    st.caption(f"Base carregada: {source_label}")
    st.caption(APP_VERSION)

# -----------------------------------------------------------------------------
# Filtros globais
# -----------------------------------------------------------------------------
with st.sidebar:
    st.divider()
    st.markdown("### Filtros")

    years = sorted([int(x) for x in df["Order Year"].dropna().unique().tolist()])
    selected_years = st.multiselect("Ano do pedido", years, default=years)

    supplier_groups = sorted(df["Supplier Group Description"].dropna().astype(str).unique().tolist())
    selected_supplier_groups = st.multiselect("Grupo fornecedor", supplier_groups, default=[])

    clerks = sorted(df["Clerk name"].dropna().astype(str).unique().tolist())
    selected_clerks = st.multiselect("Comprador", clerks, default=[])

    suppliers = sorted(df["Supplier Name"].dropna().astype(str).unique().tolist())
    selected_suppliers = st.multiselect("Fornecedor", suppliers, default=[])

    countries = sorted(df["Supplier's Country"].dropna().astype(str).unique().tolist())
    selected_countries = st.multiselect("País fornecedor", countries, default=[])

    currencies = sorted(df["Currency"].dropna().astype(str).unique().tolist())
    selected_currencies = st.multiselect("Moeda", currencies, default=[])

filtered = df.copy()
if selected_years:
    filtered = filtered[filtered["Order Year"].isin(selected_years)]
if selected_supplier_groups:
    filtered = filtered[filtered["Supplier Group Description"].isin(selected_supplier_groups)]
if selected_clerks:
    filtered = filtered[filtered["Clerk name"].isin(selected_clerks)]
if selected_suppliers:
    filtered = filtered[filtered["Supplier Name"].isin(selected_suppliers)]
if selected_countries:
    filtered = filtered[filtered["Supplier's Country"].isin(selected_countries)]
if selected_currencies:
    filtered = filtered[filtered["Currency"].isin(selected_currencies)]

if filtered.empty:
    st.warning("Nenhum registro encontrado para os filtros selecionados.")
    st.stop()

# -----------------------------------------------------------------------------
# Métricas gerais
# -----------------------------------------------------------------------------
total_spend = filtered["Spend BRL"].sum()
open_value = filtered.loc[filtered["Is Open"], "Open Value Est."].sum()
overdue_value = filtered.loc[filtered["Is Overdue"], "Open Value Est."].sum()
orders = filtered["Order No."].nunique()
suppliers_count = filtered["Supplier Name"].nunique()
items_count = filtered["Part No."].nunique()
fill_rate = filtered["Delivered Qty"].sum() / filtered["Order Qty"].sum() if filtered["Order Qty"].sum() else np.nan
pmp = (filtered["Spend BRL"] * filtered["Maturity days"].fillna(0)).sum() / total_spend if total_spend else np.nan

k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Spend total", fmt_brl(total_spend, 0), f"{fmt_num(orders, 0)} pedidos")
with k2:
    kpi_card("Carteira aberta", fmt_brl(open_value, 0), f"{fmt_num(filtered['Is Open'].sum(), 0)} linhas abertas")
with k3:
    kpi_card("Valor atrasado", fmt_brl(overdue_value, 0), "Saldo aberto com Conf. Date vencida")
with k4:
    kpi_card("PMP ponderado", f"{fmt_num(pmp, 1)} dias", "Ponderado por Pos. Value")

k5, k6, k7, k8 = st.columns(4)
with k5:
    kpi_card("Fill rate", fmt_pct(fill_rate, 1), "Pos. Del. Qua. / Order Qua.")
with k6:
    kpi_card("Fornecedores", fmt_num(suppliers_count, 0), "ativos na seleção")
with k7:
    kpi_card("Itens", fmt_num(items_count, 0), "Part No. distintos")
with k8:
    import_share = filtered.loc[filtered["Supplier's Country"] != "BRA", "Spend BRL"].sum() / total_spend if total_spend else np.nan
    kpi_card("Spend importado", fmt_pct(import_share, 1), "Supplier's Country diferente de BRA")

st.write("")

tab_, tab_follow, tab_suppliers, tab_abc_sup, tab_abc_items, tab_buyers, tab_pmp, tab_price, tab_radar = st.tabs([
    "🏠 Control Tower",
    "📦 Follow-up",
    "🏭 Fornecedores",
    "📊 ABC Fornecedores",
    "📦 ABC Itens",
    "👤 Compradores",
    "💰 PMP",
    "📈 Price Intelligence",
    "⚡ Radar",
])

# -----------------------------------------------------------------------------
# Control Tower
# -----------------------------------------------------------------------------
with tab_control:
    c1, c2 = st.columns([1.35, 1])
    with c1:
        st.plotly_chart(monthly_spend_chart(filtered), use_container_width=True)
    with c2:
        by_group = filtered.groupby("Supplier Group Description", as_index=False)["Spend BRL"].sum().sort_values("Spend BRL", ascending=False).head(12)
        fig_group = px.treemap(by_group, path=["Supplier Group Description"], values="Spend BRL", title="Spend por grupo de fornecedor")
        fig_group.update_layout(height=410, margin=dict(t=55, r=10, l=10, b=10))
        st.plotly_chart(fig_group, use_container_width=True)

    st.markdown('<div class="section-title">Visão executiva</div>', unsafe_allow_html=True)
    s1, s2 = st.columns(2)
    with s1:
        abc_sup = abc_classification(filtered, "Supplier Name")
        st.plotly_chart(pareto_chart(abc_sup, "Supplier Name", title="Pareto de fornecedores"), use_container_width=True)
    with s2:
        by_clerk = filtered.groupby("Clerk name", as_index=False).agg(Spend=("Spend BRL", "sum"), Pedidos=("Order No.", "nunique")).sort_values("Spend", ascending=False).head(15)
        fig = px.bar(by_clerk, x="Spend", y="Clerk name", orientation="h", title="Spend por comprador", text="Spend")
        fig.update_traces(marker_color="#0F766E", texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(height=440, margin=dict(t=55, r=35, l=18, b=30), yaxis={"categoryorder": "total ascending"}, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

# -----------------------------------------------------------------------------
# Follow-up
# -----------------------------------------------------------------------------
with tab_follow:
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        kpi_card("Linhas abertas", fmt_num(filtered["Is Open"].sum(), 0), "Pos. Open Qua > 0")
    with f2:
        kpi_card("Linhas atrasadas", fmt_num(filtered["Is Overdue"].sum(), 0), "Conf. Date < hoje")
    with f3:
        kpi_card("Valor aberto", fmt_brl(open_value, 0), "estimado pelo saldo")
    with f4:
        kpi_card("Valor atrasado", fmt_brl(overdue_value, 0), "estimado pelo saldo")

    open_df = filtered[filtered["Is Open"]].copy()
    if open_df.empty:
        st.info("Não há linhas abertas nos filtros atuais.")
    else:
        bins = [-1, 0, 7, 15, 30, 60, 99999]
        labels = ["Em dia", "1-7", "8-15", "16-30", "31-60", ">60"]
        open_df["Aging"] = pd.cut(open_df["Days Late"].fillna(0), bins=bins, labels=labels)
        aging = open_df.groupby("Aging", observed=False, as_index=False)["Open Value Est."].sum()
        fig_aging = px.bar(aging, x="Aging", y="Open Value Est.", title="Aging de carteira aberta", text="Open Value Est.")
        fig_aging.update_traces(marker_color="#DC2626", texttemplate="%{text:,.0f}", textposition="outside")
        fig_aging.update_layout(height=380, margin=dict(t=55, r=18, l=18, b=35), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig_aging, use_container_width=True)

        follow_cols = ["Order No.", "Pos. No.", "Supplier Name", "Part No.", "Name", "Clerk name", "Desired Date", "Conf. Date", "Delivery Date", "Order Qty", "Delivered Qty", "Open Qty", "Open Value Est.", "Days Late"]
        follow_table = open_df[follow_cols].sort_values(["Is Overdue", "Days Late", "Open Value Est."], ascending=[False, False, False]).copy()
        st.dataframe(follow_table, use_container_width=True, height=520)
        st.download_button("Exportar follow-up", export_excel_bytes(follow_table), "followup_compras.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------------------------------------------------------------
# Fornecedores
# -----------------------------------------------------------------------------
with tab_suppliers:
    supplier_score = filtered.groupby(["Supplier Group Description", "Supplier Name"], as_index=False).agg(
        Spend=("Spend BRL", "sum"),
        Pedidos=("Order No.", "nunique"),
        Linhas=("Order No.", "count"),
        Valor_Aberto=("Open Value Est.", "sum"),
        Maturidade=("Maturity days", "mean"),
        Fill_Rate=("Fill Rate", "mean"),
    ).sort_values("Spend", ascending=False)
    supplier_score["PMP"] = supplier_score["Maturidade"]
    c1, c2 = st.columns([1, 1])
    with c1:
        grp = supplier_score.groupby("Supplier Group Description", as_index=False).agg(Spend=("Spend", "sum")).sort_values("Spend", ascending=False)
        fig = px.bar(grp.head(15), x="Spend", y="Supplier Group Description", orientation="h", title="Spend por grupo de fornecedor")
        fig.update_layout(height=460, yaxis={"categoryorder": "total ascending"}, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.scatter(supplier_score, x="Pedidos", y="Spend", size="Valor_Aberto", color="Supplier Group Description", hover_name="Supplier Name", title="Matriz fornecedor: pedidos x spend x backlog")
        fig.update_layout(height=460, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(format_table_currency(supplier_score, currency_cols=["Spend", "Valor_Aberto"], num_cols=["Pedidos", "Linhas"]), use_container_width=True, height=460)

# -----------------------------------------------------------------------------
# ABC fornecedores
# -----------------------------------------------------------------------------
with tab_abc_sup:
    abc_sup = abc_classification(filtered, "Supplier Name")
    st.plotly_chart(pareto_chart(abc_sup, "Supplier Name", title="ABC fornecedores | Pareto por Pos. Value"), use_container_width=True)
    st.dataframe(format_table_currency(abc_sup, currency_cols=["Spend BRL"], pct_cols=["Share %", "Cum %"], num_cols=["Pedidos", "Linhas"]), use_container_width=True, height=520)

# -----------------------------------------------------------------------------
# ABC itens
# -----------------------------------------------------------------------------
with tab_abc_items:
    item_base = filtered.groupby(["Part No.", "Name"], dropna=False, as_index=False).agg(
        **{"Spend BRL": ("Spend BRL", "sum"), "Pedidos": ("Order No.", "nunique"), "Linhas": ("Order No.", "count"), "Preço Médio HC": ("Price H. C.", "mean")}
    ).sort_values("Spend BRL", ascending=False).reset_index(drop=True)
    total_items_spend = item_base["Spend BRL"].sum()
    item_base["Share %"] = np.where(total_items_spend > 0, item_base["Spend BRL"] / total_items_spend, 0)
    item_base["Cum %"] = item_base["Share %"].cumsum()
    item_base["ABC"] = np.select([item_base["Cum %"] <= 0.80, item_base["Cum %"] <= 0.95], ["A", "B"], default="C")
    st.plotly_chart(pareto_chart(item_base.rename(columns={"Part No.": "Item"}), "Item", title="ABC itens | Pareto por Pos. Value"), use_container_width=True)
    st.dataframe(format_table_currency(item_base, currency_cols=["Spend BRL"], pct_cols=["Share %", "Cum %"], num_cols=["Pedidos", "Linhas"]), use_container_width=True, height=520)

# -----------------------------------------------------------------------------
# Compradores
# -----------------------------------------------------------------------------
with tab_buyers:
    buyers = filtered.groupby("Clerk name", as_index=False).agg(
        Spend=("Spend BRL", "sum"),
        Pedidos=("Order No.", "nunique"),
        Linhas=("Order No.", "count"),
        Fornecedores=("Supplier Name", "nunique"),
        Fill_Rate=("Fill Rate", "mean"),
        PMP=("Maturity days", "mean"),
        Backlog=("Open Value Est.", "sum"),
    ).sort_values("Spend", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(buyers, x="Spend", y="Clerk name", orientation="h", title="Spend por comprador", text="Spend")
        fig.update_layout(height=460, yaxis={"categoryorder": "total ascending"}, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.scatter(buyers, x="Pedidos", y="Spend", size="Fornecedores", color="Fill_Rate", hover_name="Clerk name", title="Matriz comprador: pedidos x spend x fornecedores")
        fig.update_layout(height=460, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(format_table_currency(buyers, currency_cols=["Spend", "Backlog"], num_cols=["Pedidos", "Linhas", "Fornecedores"]), use_container_width=True, height=460)

# -----------------------------------------------------------------------------
# PMP
# -----------------------------------------------------------------------------
with tab_pmp:
    st.markdown('<div class="section-title">Prazo médio de pagamento ponderado</div>', unsafe_allow_html=True)
    pmp_supplier = filtered.dropna(subset=["Maturity days"]).groupby("Supplier Name", as_index=False).apply(
        lambda g: pd.Series({
            "Spend": g["Spend BRL"].sum(),
            "PMP": (g["Spend BRL"] * g["Maturity days"].fillna(0)).sum() / g["Spend BRL"].sum() if g["Spend BRL"].sum() else np.nan,
            "Pedidos": g["Order No."].nunique(),
        })
    ).reset_index(drop=True).sort_values("Spend", ascending=False)
    pmp_group = filtered.dropna(subset=["Maturity days"]).groupby("Supplier Group Description", as_index=False).apply(
        lambda g: pd.Series({
            "Spend": g["Spend BRL"].sum(),
            "PMP": (g["Spend BRL"] * g["Maturity days"].fillna(0)).sum() / g["Spend BRL"].sum() if g["Spend BRL"].sum() else np.nan,
        })
    ).reset_index(drop=True).sort_values("Spend", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(pmp_supplier.head(20), x="PMP", y="Supplier Name", orientation="h", title="PMP por fornecedor | Top spend")
        fig.update_layout(height=520, yaxis={"categoryorder": "total ascending"}, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(pmp_group.head(20), x="PMP", y="Supplier Group Description", orientation="h", title="PMP por grupo de fornecedor")
        fig.update_layout(height=520, yaxis={"categoryorder": "total ascending"}, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(format_table_currency(pmp_supplier, currency_cols=["Spend"], num_cols=["Pedidos"]), use_container_width=True, height=420)

# -----------------------------------------------------------------------------
# Price Intelligence
# -----------------------------------------------------------------------------
with tab_price:
    st.markdown('<div class="section-title">Histórico de preços sem efeito câmbio</div>', unsafe_allow_html=True)
    st.markdown('<div class="small-muted">A análise usa Price H. C., ou seja, preço do item em BRL.</div>', unsafe_allow_html=True)
    valid_price = filtered.dropna(subset=["Order Date", "Price H. C."]).copy()
    if valid_price.empty:
        st.info("Sem dados válidos de preço para os filtros atuais.")
    else:
        options = sorted(valid_price["Part No."].astype(str).unique().tolist())
        selected_part = st.selectbox("Selecione o Part No.", options)
        part_df = valid_price[valid_price["Part No."].astype(str) == selected_part].sort_values("Order Date")
        last_price = part_df["Price H. C."].iloc[-1]
        avg_price = part_df["Price H. C."].mean()
        min_price = part_df["Price H. C."].min()
        max_price = part_df["Price H. C."].max()
        p1, p2, p3, p4 = st.columns(4)
        with p1: kpi_card("Último preço", fmt_brl(last_price, 2), selected_part)
        with p2: kpi_card("Média histórica", fmt_brl(avg_price, 2), "Price H. C.")
        with p3: kpi_card("Mínimo histórico", fmt_brl(min_price, 2), "Price H. C.")
        with p4: kpi_card("Máximo histórico", fmt_brl(max_price, 2), "Price H. C.")
        fig = px.line(part_df, x="Order Date", y="Price H. C.", color="Supplier Name", markers=True, title=f"Histórico de preço | {selected_part}")
        fig.update_layout(height=480, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

        changes = price_changes(filtered)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### Maiores aumentos")
            st.dataframe(changes.head(20), use_container_width=True, height=360)
        with c2:
            st.markdown("##### Maiores reduções")
            st.dataframe(changes.tail(20).sort_values("Variação %"), use_container_width=True, height=360)

# -----------------------------------------------------------------------------
# Radar
# -----------------------------------------------------------------------------
with tab_radar:
    r1, r2 = st.columns(2)
    open_rank = filtered[filtered["Is Open"]].groupby("Supplier Name", as_index=False)["Open Value Est."].sum().sort_values("Open Value Est.", ascending=False).head(15)
    late_rank = filtered[filtered["Is Overdue"]].groupby("Supplier Name", as_index=False).agg(Valor_Atrasado=("Open Value Est.", "sum"), Dias_Medios=("Days Late", "mean")).sort_values("Valor_Atrasado", ascending=False).head(15)
    with r1:
        st.markdown("##### Top backlog aberto")
        st.dataframe(format_table_currency(open_rank, currency_cols=["Open Value Est."]), use_container_width=True, height=420)
    with r2:
        st.markdown("##### Top atrasos")
        st.dataframe(format_table_currency(late_rank, currency_cols=["Valor_Atrasado"]), use_container_width=True, height=420)

    changes = price_changes(filtered)
    if not changes.empty:
        r3, r4 = st.columns(2)
        with r3:
            st.markdown("##### Top reajustes de preço")
            st.dataframe(changes.head(15), use_container_width=True, height=420)
        with r4:
            st.markdown("##### Top reduções de preço")
            st.dataframe(changes.tail(15).sort_values("Variação %"), use_container_width=True, height=420)

st.write("")
st.caption(APP_VERSION)
