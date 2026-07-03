
import os
import re
from datetime import datetime
from io import BytesIO
from html import escape

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

EXCEL_FILE = "app_paperbase.xlsx"
SOURCE_SHEET = "PAPER BASE"
APP_TITLE = "Paper Base Dashboard"
APP_VERSION = "V1.2 | Paper Base"
TABLE_HEIGHT_PX = 900
LOGO_PATH = "assets/impress_logo.png"

# Logo nativa: posicionada acima do menu multipágina do Streamlit.
if os.path.exists(LOGO_PATH):
    st.logo(LOGO_PATH, size="large")

st.markdown(
    """
    <style>
        .block-container {padding-top: 1.1rem; padding-bottom: 1.2rem;}
        .pb-page-title {font-size:1.75rem; font-weight:700; color:#111827; margin:0 0 .15rem 0;}
        .pb-subtitle {font-size:.92rem; color:#6B7280; margin-bottom:1rem;}
        .pb-kpi {border:1px solid #E5E7EB; border-radius:14px; padding:14px 16px; background:#FFF; box-shadow:0 1px 2px rgba(0,0,0,.04); min-height:82px;}
        .pb-kpi-label {font-size:.78rem; color:#6B7280; margin-bottom:6px;}
        .pb-kpi-value {font-size:1.16rem; font-weight:700; color:#111827; line-height:1.2;}
        .pb-table-shell {border:1px solid #E5E7EB; border-radius:14px; background:#FFF; overflow:hidden; box-shadow:0 1px 2px rgba(0,0,0,.03);}
        .pb-table-wrap {overflow-y:auto; overflow-x:auto;}
        .pb-table-wrap.no-scroll {overflow:visible;}
        .pb-table {width:100%; border-collapse:separate; border-spacing:0; table-layout:auto; font-size:.84rem; color:#111827;}
        .pb-table thead th {position:sticky; top:0; z-index:2; background:#F9FAFB; color:#6B7280; font-weight:600; border-bottom:1px solid #E5E7EB; border-right:1px solid #F0F2F5; padding:8px 10px; white-space:nowrap;}
        .pb-table tbody td {padding:7px 10px; border-bottom:1px solid #F3F4F6; border-right:1px solid #F8FAFC; white-space:nowrap; background:#FFF;}
        .pb-table tbody tr:hover td {background:#FAFBFF;}
        .pb-left{text-align:left}.pb-center{text-align:center}.pb-right{text-align:right;font-variant-numeric:tabular-nums}.pb-note{color:#6B7280;font-size:.78rem;margin-top:.35rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div style='font-size:0.92rem;color:#111827;'><b>app_paperbase.xlsx</b></div>
        <div style='font-size:0.88rem;color:#4B5563;'>Última atualização: {excel_last_update}</div>
        """,
        unsafe_allow_html=True,
    )
    st.caption("v1.0")

def get_excel_last_update(file_path):
    try:
        wb = load_workbook(file_path, read_only=True, keep_vba=True)
        modified = wb.properties.modified
        if modified is not None:
            if hasattr(modified, "tzinfo") and modified.tzinfo is not None:
                modified = modified.replace(tzinfo=None)
            return modified.strftime("%d/%m/%Y")
    except Exception:
        pass
    try:
        ts = os.path.getmtime(file_path)
        dt = datetime.fromtimestamp(ts)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return "N/A"
        
def parse_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("€", "").replace(" ", "")
    s = re.sub(r"[^0-9,.-]", "", s)
    if s in {"", "-", ".", ","}:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def format_br_number(value, decimals=2):
    if pd.isna(value):
        return ""
    try:
        n = float(value)
    except Exception:
        return str(value)
    return f"{n:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def format_no_decimal(value):
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return str(value)


def supplier_code(value):
    return "" if pd.isna(value) else str(value).strip()[:6].strip()


def supplier_clean(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return text.split(" / ", 1)[1].strip() if " / " in text else (text[9:].strip() if len(text) >= 10 else text)


def parse_excel_month(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value.year, value.month, 1)
    n = parse_number(value)
    if n is not None:
        dt = pd.to_datetime(n, unit="D", origin="1899-12-30", errors="coerce")
        if pd.notna(dt):
            return pd.Timestamp(dt.year, dt.month, 1)
    dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return pd.Timestamp(dt.year, dt.month, 1) if pd.notna(dt) else pd.NaT


def kpi_card(label, value):
    st.markdown(
        f"""
        <div class='pb-kpi'>
            <div class='pb-kpi-label'>{escape(str(label))}</div>
            <div class='pb-kpi-value'>{escape(str(value))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Paper Base")
    output.seek(0)
    return output.getvalue()


def build_html_table(df_html, height_px=430, no_scroll=False):
    if df_html.empty:
        return "<div class='pb-note'>Sem registros para exibir.</div>"

    right_cols = {"Quantity KG", "Value EUR", "Price Avg", "Share %", "EUR/kg"}
    left_cols = {"Supplier", "Mat Description", "Material Group"}

    wrap_class = "pb-table-wrap no-scroll" if no_scroll else "pb-table-wrap"
    max_height = "" if no_scroll else f" style='max-height:{int(height_px)}px;'"

    parts = [
        "<div class='pb-table-shell'>",
        f"<div class='{wrap_class}'{max_height}>",
        "<table class='pb-table'><thead><tr>",
    ]

    for col in df_html.columns:
        cls = "pb-right" if col in right_cols else ("pb-left" if col in left_cols else "pb-center")
        parts.append(f"<th class='{cls}'>{escape(str(col))}</th>")

    parts.append("</tr></thead><tbody>")

    for _, row in df_html.iterrows():
        parts.append("<tr>")
        for col in df_html.columns:
            cls = "pb-right" if col in right_cols else ("pb-left" if col in left_cols else "pb-center")
            value = "" if pd.isna(row[col]) else str(row[col])
            parts.append(f"<td class='{cls}'>{escape(value)}</td>")
        parts.append("</tr>")

    parts.append("</tbody></table></div></div>")
    return "".join(parts)


@st.cache_data(show_spinner=False)
def load_data():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, engine="openpyxl")
    raw.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in raw.columns]

    required = ["Supplier name", "Mat Description", "Quantity KG", "Value EUR", "Month.Year", "Material Group"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias não encontradas: {missing}")

    df = raw.copy()
    df["Supplier Raw"] = df["Supplier name"].astype(str).str.strip()
    df["Supplier Code"] = df["Supplier Raw"].apply(supplier_code)
    df["Supplier"] = df["Supplier Raw"].apply(supplier_clean)
    df["Mat Description"] = df["Mat Description"].astype(str).str.strip()
    df["Material Group"] = df["Material Group"].astype(str).str.strip()
    df["Quantity KG"] = df["Quantity KG"].apply(parse_number)
    df["Value EUR"] = df["Value EUR"].apply(parse_number)
    df["Month Date"] = df["Month.Year"].apply(parse_excel_month)
    df = df.dropna(subset=["Supplier", "Quantity KG", "Value EUR", "Month Date"])
    df = df[(df["Quantity KG"] > 0) & (df["Value EUR"] > 0)]
    df["Year"] = df["Month Date"].dt.year.astype(int)
    df["Month"] = df["Month Date"].dt.month.astype(int)
    df["EUR/kg"] = df["Value EUR"] / df["Quantity KG"]
    return df.reset_index(drop=True)


def supplier_summary(df_in):
    """Ranking por fornecedor, sem coluna Rank e com ordem profissional de colunas."""
    ordered_cols = ["Supplier", "Quantity KG", "Value EUR", "Price Avg", "Share %"]
    if df_in.empty:
        return pd.DataFrame(columns=ordered_cols)

    base = df_in.groupby("Supplier", as_index=False).agg(
        **{"Quantity KG": ("Quantity KG", "sum"), "Value EUR": ("Value EUR", "sum")}
    )
    base["Price Avg"] = base["Value EUR"] / base["Quantity KG"]
    total_kg = base["Quantity KG"].sum()
    base["Share %"] = (base["Quantity KG"] / total_kg * 100) if total_kg else 0
    base = base.sort_values("Quantity KG", ascending=False, kind="stable").reset_index(drop=True)
    return base[ordered_cols]


def format_supplier_summary(df_in):
    df = df_in.copy()
    if "Quantity KG" in df.columns:
        df["Quantity KG"] = df["Quantity KG"].apply(format_no_decimal)
    if "Value EUR" in df.columns:
        df["Value EUR"] = df["Value EUR"].apply(lambda x: format_br_number(x, 0))
    if "Price Avg" in df.columns:
        df["Price Avg"] = df["Price Avg"].apply(lambda x: format_br_number(x, 2))
    if "Share %" in df.columns:
        df["Share %"] = df["Share %"].apply(lambda x: format_br_number(x, 1))
    return df


def monthly_summary(df_in):
    if df_in.empty:
        return pd.DataFrame(columns=["Year", "Month", "Quantity KG", "Quantity ton", "Value EUR", "EUR/kg"])
    base = df_in.groupby(["Year", "Month"], as_index=False).agg(
        **{"Quantity KG": ("Quantity KG", "sum"), "Value EUR": ("Value EUR", "sum")}
    )
    base["EUR/kg"] = base["Value EUR"] / base["Quantity KG"]
    base["Quantity ton"] = base["Quantity KG"] / 1000
    return base.sort_values(["Year", "Month"]).reset_index(drop=True)

def build_monthly_chart(monthly_df):
    names = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}
    months = list(range(1, 13))
    x = [names[m] for m in months]
    years = sorted(monthly_df["Year"].unique().tolist()) if not monthly_df.empty else []

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Barras: mantemos tons fortes e parecidos com o Excel.
    bar_palette = ["#ED7D31", "#70AD47", "#A5A5A5", "#FFC000"]

    # Linhas de preço: cores mais contrastantes para melhorar leitura sobre as barras.
    line_palette = ["#5B9BD5", "#ED7D31", "#70AD47", "#7F7F7F"]

    for idx, year in enumerate(years):
        ydf = monthly_df[monthly_df["Year"] == year].set_index("Month")
        qty = [ydf.loc[m, "Quantity ton"] if m in ydf.index else 0 for m in months]
        avg = [ydf.loc[m, "EUR/kg"] if m in ydf.index else None for m in months]

        bar_color = bar_palette[idx % len(bar_palette)]
        line_color = line_palette[idx % len(line_palette)]

        # Quantidade: label dentro da barra, ancorado na base.
        # Isso deixa o topo da coluna livre para a linha/labels de Price Avg.
        fig.add_trace(
            go.Bar(
                x=x,
                y=qty,
                name=f"{year} ton",
                marker_color=bar_color,
                marker_line=dict(color="rgba(255,255,255,0.85)", width=0.5),
                text=[format_no_decimal(v) if v else "" for v in qty],
                textposition="inside",
                insidetextanchor="start",
                textfont=dict(size=11, color="#1F2937"),
                hovertemplate="Mês: %{x}<br>Quantidade: %{y:.0f} ton<extra></extra>",
            ),
            secondary_y=False,
        )

        # Price Avg: linha sem texto nativo. Os labels serão annotations com fundo.
        fig.add_trace(
            go.Scatter(
                x=x,
                y=avg,
                name=f"Price Avg {year}",
                mode="lines+markers",
                line=dict(color=line_color, width=2.4, dash="dot"),
                marker=dict(size=7, color=line_color, line=dict(color="white", width=1.2)),
                hovertemplate="Mês: %{x}<br>Price Avg: %{y:.2f} EUR/kg<extra></extra>",
                connectgaps=False,
            ),
            secondary_y=True,
        )

        # Labels com fundo branco/translúcido e borda sutil.
        # Isso melhora muito a leitura quando o texto cruza barras ou gridlines.
        for month_label, price in zip(x, avg):
            if price is None or pd.isna(price):
                continue
            fig.add_annotation(
                x=month_label,
                y=price,
                xref="x",
                yref="y2",
                text=format_br_number(price, 2),
                showarrow=False,
                yshift=16,
                font=dict(size=11, color=line_color),
                bgcolor="rgba(255,255,255,0.88)",
                bordercolor="rgba(107,114,128,0.28)",
                borderwidth=1,
                borderpad=2,
            )

    fig.update_layout(
        title=dict(text="Monthly Imports | Base Paper", x=.5, xanchor="center", font=dict(size=16, color="#111827")),
        height=430,
        barmode="group",
        margin=dict(t=60, r=58, l=54, b=48),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", y=-.16, x=.5, xanchor="center"),
        bargap=0.32,
        bargroupgap=0.08,
    )

    fig.update_yaxes(
        title_text="Toneladas",
        secondary_y=False,
        showgrid=True,
        gridcolor="#E5E7EB",
        zeroline=True,
        zerolinecolor="#D1D5DB",
    )
    fig.update_yaxes(
        title_text="EUR/kg",
        secondary_y=True,
        tickformat=".2f",
        showgrid=False,
    )
    fig.update_xaxes(showgrid=False)
    return fig
try:
    df = load_data()
except Exception as e:
    st.error("❌ Falha ao ler a aba PAPER BASE do arquivo app_paperbase.xlsx.")
    st.code(str(e))
    st.stop()

st.markdown(f"<div class='pb-page-title'>{APP_TITLE}</div>", unsafe_allow_html=True)
st.markdown("<div class='pb-subtitle'>Fornecimento mensal, ranking por fornecedor e custo médio ponderado em EUR/kg.</div>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Filtros")
    st.caption(APP_VERSION)
    st.divider()
    years = sorted(df["Year"].unique().tolist())
    selected_years = st.multiselect("Ano", years, default=years)
    materials = sorted(df["Material Group"].dropna().unique().tolist())
    selected_materials = st.multiselect("Material Group", materials, default=materials)
    suppliers = sorted(df["Supplier"].dropna().unique().tolist())
    selected_suppliers = st.multiselect("Fornecedor", suppliers, default=[])
    month_map = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}
    month_options = sorted(df["Month"].unique().tolist())
    selected_month_labels = st.multiselect("Meses", [month_map[m] for m in month_options], default=[month_map[m] for m in month_options])
    selected_months = [m for m, lbl in month_map.items() if lbl in selected_month_labels]
    st.caption("Fonte: app_paperbase.xlsx | aba PAPER BASE")

filtered = df.copy()
if selected_years:
    filtered = filtered[filtered["Year"].isin(selected_years)]
if selected_materials:
    filtered = filtered[filtered["Material Group"].isin(selected_materials)]
if selected_suppliers:
    filtered = filtered[filtered["Supplier"].isin(selected_suppliers)]
if selected_months:
    filtered = filtered[filtered["Month"].isin(selected_months)]

if filtered.empty:
    st.warning("Nenhum registro encontrado para os filtros selecionados.")
    st.stop()

latest = filtered["Month Date"].max()
ytd_year = int(latest.year)
last_df = filtered[filtered["Month Date"] == latest]
ytd_df = filtered[(filtered["Year"] == ytd_year) & (filtered["Month Date"] <= latest)]

total_kg = filtered["Quantity KG"].sum()
total_eur = filtered["Value EUR"].sum()
avg = total_eur / total_kg if total_kg else None

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    kpi_card("Quantidade total", f"{format_no_decimal(total_kg / 1000)} ton")
with k2:
    kpi_card("Valor total", f"€ {format_br_number(total_eur, 0)}")
with k3:
    kpi_card("Custo médio ponderado", f"€ {format_br_number(avg, 2)}/kg")
with k4:
    kpi_card("Fornecedores", format_no_decimal(filtered["Supplier"].nunique()))
with k5:
    kpi_card("Materiais", format_no_decimal(filtered["Mat Description"].nunique()))

st.markdown("### Monthly Imports | Base Paper")
st.plotly_chart(build_monthly_chart(monthly_summary(filtered)), width="stretch")
st.markdown("<div class='pb-note'>Barras: toneladas, com label na base da coluna. Linha: Price Avg com fundo nos labels para melhorar a leitura.</div>", unsafe_allow_html=True)

st.markdown("### Ranking por fornecedor")
c1, c2 = st.columns(2)
with c1:
    st.markdown(f"#### Último mês: {latest.strftime('%m/%Y')}")
    st.markdown(build_html_table(format_supplier_summary(supplier_summary(last_df)), TABLE_HEIGHT_PX, no_scroll=True), unsafe_allow_html=True)
with c2:
    st.markdown(f"#### YTD: {ytd_year}")
    st.markdown(build_html_table(format_supplier_summary(supplier_summary(ytd_df)), TABLE_HEIGHT_PX, no_scroll=True), unsafe_allow_html=True)

with st.expander("Ver base tratada / exportar"):
    detail = filtered[["Year", "Month", "Supplier", "Mat Description", "Quantity KG", "Value EUR", "EUR/kg", "Material Group"]].copy().sort_values(["Year", "Month", "Supplier", "Mat Description"])
    display = detail.copy()
    display["Quantity KG"] = display["Quantity KG"].apply(format_no_decimal)
    display["Value EUR"] = display["Value EUR"].apply(lambda x: format_br_number(x, 2))
    display["EUR/kg"] = display["EUR/kg"].apply(lambda x: format_br_number(x, 2))
    st.markdown(build_html_table(display, 420), unsafe_allow_html=True)
    e1, e2, _ = st.columns([1.2, 1.2, 6])
    with e1:
        st.download_button("Exportar CSV", detail.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig"), "paperbase_filtrado.csv", "text/csv", width="stretch")
    with e2:
        st.download_button("Exportar Excel", to_excel_bytes(detail), "paperbase_filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

st.caption("Regra: Price Avg = soma de Value EUR dividida pela soma de Quantity KG no período selecionado.")
