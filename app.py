
import re
import unicodedata
from io import BytesIO
from datetime import datetime
import os

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Cockpit Papel", layout="wide")

EXCEL_FILE = "Cockpit_Papel.xlsm"
SOURCE_SHEET = "Preços e Condições"
PREMISSAS_SHEET = "Premissas"
APP_TITLE = "Cockpit Papel"

st.title(APP_TITLE)


def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    text = text.replace("
", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def parse_number(value):
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none", "-", "--"}:
        return None

    s = s.replace("R$", "").replace("$", "").replace("€", "")
    s = s.replace("%", "")
    s = s.replace(" ", "")
    s = re.sub(r"[^0-9,.-]", "", s)

    if s in {"", "-", ".", ","}:
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        if s.count(".") > 1:
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(s)
    except Exception:
        return None


def series_to_numeric(series):
    return series.apply(parse_number)


def format_br_number(value, decimals=2):
    if pd.isna(value):
        return ""
    try:
        n = float(value)
    except Exception:
        return str(value)
    s = f"{n:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def format_no_decimal(value):
    if pd.isna(value):
        return ""
    try:
        return str(int(round(float(value))))
    except Exception:
        return str(value)


def safe_display_string(value, numeric_no_decimal=False):
    if pd.isna(value):
        return ""
    if numeric_no_decimal:
        return format_no_decimal(value)
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return ""
        return value.strftime("%d/%m/%Y")
    return str(value)


def detect_header_row(raw_df):
    targets = ["impress type", "supplier", "current price"]
    best_row = None
    best_score = -1
    max_rows = min(len(raw_df), 50)

    for i in range(max_rows):
        row_values = [normalize_text(v) for v in raw_df.iloc[i].tolist()]
        row_text = " | ".join(row_values)
        score = sum(1 for t in targets if t in row_text)
        if score > best_score:
            best_score = score
            best_row = i

    return best_row if best_score >= 2 else None


def find_column(columns, aliases):
    normalized_map = {col: normalize_text(col) for col in columns}
    for alias in aliases:
        alias_norm = normalize_text(alias)
        for col, col_norm in normalized_map.items():
            if col_norm == alias_norm:
                return col
        for col, col_norm in normalized_map.items():
            if alias_norm in col_norm:
                return col
    return None


def deduplicate_for_analysis(df_in):
    """
    Regras implementadas:
    - Material Number e Width (mm) NÃO entram na visão analítica.
    - Se após remover esses campos sobrarem linhas idênticas,
      mantém apenas a linha com a data mais recente em 'Última Atualização de Preço'.
    """
    df_work = df_in.copy()

    if df_work.empty:
        return df_work

    date_col = "Última Atualização de Preço"
    if date_col in df_work.columns:
        df_work[date_col] = pd.to_datetime(df_work[date_col], errors="coerce", dayfirst=True)
        df_work = df_work.sort_values(by=date_col, ascending=False, na_position="last", kind="stable")

    key_cols = [c for c in df_work.columns if c != date_col]
    if key_cols:
        df_work = df_work.drop_duplicates(subset=key_cols, keep="first")

    final_sort = [c for c in ["Impress Type", "g/m2", "Supplier", "Lot (ton)"] if c in df_work.columns]
    if final_sort:
        df_work = df_work.sort_values(final_sort, kind="stable")

    return df_work.reset_index(drop=True)


def build_canonical_dataframe(df):
    aliases = {
        "Material Number": ["Material Number", "Material", "Código Material", "Codigo Material", "Item Code"],
        "Impress Type": ["Impress Type", "Print Type"],
        "Width (mm)": ["Width (mm)", "Width mm", "Width", "Largura", "Largura (mm)"],
        "g/m2": ["g/m2", "g/m²", "Gramatura", "gsm"],
        "Supplier": ["Supplier", "Fornecedor"],
        "Currency": ["Currency", "Moeda"],
        "Current Price": ["Current Price", "Preço Atual", "Preco Atual", "CurrentPrice"],
        "Paper bonus (t)": ["Paper bonus (t)", "Paper bonus", "Bonus", "Bonus (t)"],
        "Lot (ton)": ["Lot (ton)", "Lot", "Lote", "Lote (ton)"],
        "TCO (R$/KG)": ["TCO (R$/KG)", "TCO R$/KG", "TCO KG", "TCO"],
        "TCO (R$/M2)": ["TCO (R$/M2)", "TCO R$/M2", "TCO M2"],
        "Payment Terms": ["Payment Terms", "Payment Term", "Prazo Pagamento"],
        "Working days": ["Working days", "Dias uteis", "Dias úteis"],
        "P.Value (R$/KG)": ["P.Value (R$/KG)", "P.Value R$/KG", "P Value (R$/KG)", "Present Value", "PV KG", "P.Value"],
        "P.Value (R$/M2)": ["P.Value (R$/M2)", "P.Value R$/M2", "P Value (R$/M2)", "PV M2"],
        "Última Atualização de Preço": ["Última Atualização de Preço", "Ultima Atualizacao de Preco", "Last Price Update", "Last Update"],
    }

    rename_map = {}
    for canonical_name, alias_list in aliases.items():
        original = find_column(df.columns, alias_list)
        if original:
            rename_map[original] = canonical_name

    df2 = df.rename(columns=rename_map).copy()

    # Mantemos Material Number e Width apenas para conseguir descartá-los conscientemente na etapa analítica.
    ordered_cols = [
        "Material Number",
        "Impress Type",
        "Width (mm)",
        "g/m2",
        "Supplier",
        "Currency",
        "Current Price",
        "Paper bonus (t)",
        "Lot (ton)",
        "TCO (R$/KG)",
        "TCO (R$/M2)",
        "Payment Terms",
        "Working days",
        "P.Value (R$/KG)",
        "P.Value (R$/M2)",
        "Última Atualização de Preço",
    ]
    existing_cols = [c for c in ordered_cols if c in df2.columns]
    df2 = df2[existing_cols].copy()

    numeric_cols = [
        "Width (mm)",
        "g/m2",
        "Current Price",
        "Paper bonus (t)",
        "Lot (ton)",
        "TCO (R$/KG)",
        "TCO (R$/M2)",
        "Working days",
        "P.Value (R$/KG)",
        "P.Value (R$/M2)",
    ]
    for col in numeric_cols:
        if col in df2.columns:
            df2[col] = series_to_numeric(df2[col])

    if "Última Atualização de Preço" in df2.columns:
        df2["Última Atualização de Preço"] = pd.to_datetime(
            df2["Última Atualização de Preço"], errors="coerce", dayfirst=True
        )

    if "TCO (R$/KG)" not in df2.columns and "Current Price" in df2.columns:
        df2["TCO (R$/KG)"] = df2["Current Price"]

    if "P.Value (R$/KG)" not in df2.columns and "TCO (R$/KG)" in df2.columns:
        df2["P.Value (R$/KG)"] = df2["TCO (R$/KG)"]

    if "TCO (R$/M2)" not in df2.columns and {"TCO (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["TCO (R$/M2)"] = df2["TCO (R$/KG)"] * (df2["g/m2"] / 1000.0)

    if "P.Value (R$/M2)" not in df2.columns and {"P.Value (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["P.Value (R$/M2)"] = df2["P.Value (R$/KG)"] * (df2["g/m2"] / 1000.0)

    essential = []
    if "Impress Type" in df2.columns:
        essential.append("Impress Type")
    if "Supplier" in df2.columns:
        essential.append("Supplier")
    if essential:
        df2 = df2.dropna(subset=essential)

    if "Current Price" in df2.columns:
        df2 = df2[df2["Current Price"].notna()]
    elif "TCO (R$/KG)" in df2.columns:
        df2 = df2[df2["TCO (R$/KG)"].notna()]

    # Remove da camada analítica os campos não relevantes.
    cols_to_drop = [c for c in ["Material Number", "Width (mm)"] if c in df2.columns]
    if cols_to_drop:
        df2 = df2.drop(columns=cols_to_drop)

    # Deduplicação: mantém a linha com a data mais recente.
    df2 = deduplicate_for_analysis(df2)

    return df2


def create_safe_multiselect(df_in, column_name, label, numeric_no_decimal=False):
    if column_name not in df_in.columns:
        return df_in

    source = df_in[column_name].dropna().copy()
    if source.empty:
        return df_in

    opt_df = pd.DataFrame({"original": source})
    opt_df["display"] = opt_df["original"].apply(
        lambda x: safe_display_string(x, numeric_no_decimal=numeric_no_decimal)
    )
    opt_df = opt_df.drop_duplicates(subset=["display"], keep="first")

    if numeric_no_decimal:
        opt_df["sort_key"] = opt_df["original"].apply(
            lambda x: parse_number(x) if parse_number(x) is not None else 10**18
        )
        opt_df = opt_df.sort_values(["sort_key", "display"], kind="stable")
    else:
        opt_df["sort_key"] = opt_df["display"].astype(str)
        opt_df = opt_df.sort_values("sort_key", kind="stable")

    display_options = opt_df["display"].tolist()
    selected_display = st.sidebar.multiselect(label, options=display_options, default=[])

    if not selected_display:
        return df_in

    selected_original = opt_df.loc[
        opt_df["display"].isin(selected_display), "original"
    ].tolist()
    return df_in[df_in[column_name].isin(selected_original)]


def safe_min(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    return s.min() if not s.empty else None


def safe_best_row(df_in, metric_col):
    if metric_col not in df_in.columns:
        return None
    tmp = df_in.copy()
    tmp[metric_col] = pd.to_numeric(tmp[metric_col], errors="coerce")
    tmp = tmp.dropna(subset=[metric_col])
    if tmp.empty:
        return None
    idx = tmp[metric_col].idxmin()
    return tmp.loc[idx]


def kpi_card(label, value):
    st.markdown(
        f"""
        <div style='padding:14px 16px;border:1px solid #E5E7EB;border-radius:12px;background:#FFFFFF;'>
            <div style='font-size:0.85rem;color:#6B7280;margin-bottom:6px;'>{label}</div>
            <div style='font-size:1.35rem;font-weight:700;color:#111827;'>{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Cockpit Filtrado")
    output.seek(0)
    return output.getvalue()


def load_premissas():
    premissas = pd.read_excel(EXCEL_FILE, sheet_name=PREMISSAS_SHEET, header=None)

    def get_cell(row_idx, col_idx):
        try:
            value = premissas.iat[row_idx, col_idx]
            return None if pd.isna(value) else value
        except Exception:
            return None

    return {
        "Frete CN": get_cell(5, 2),
        "Frete EU": get_cell(4, 2),
        "USD/BRL": get_cell(23, 2),
        "EUR/BRL": get_cell(24, 2),
        "CNY/BRL": get_cell(25, 2),
    }


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


@st.cache_data
def load_data():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=None)
    header_row = detect_header_row(raw)
    if header_row is None:
        raise ValueError(
            "Não foi possível identificar automaticamente a linha de cabeçalho da aba 'Preços e Condições'."
        )

    df = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=header_row)
    df.columns = [str(c).strip().replace("
", " ") for c in df.columns]
    df.columns = [re.sub(r"\s+", " ", c) for c in df.columns]
    df = df.dropna(how="all")
    return build_canonical_dataframe(df)


try:
    df = load_data()
except Exception as e:
    st.error("❌ Falha ao ler e estruturar a base de dados.")
    st.code(str(e))
    st.stop()

try:
    premissas_kpis = load_premissas()
except Exception as e:
    st.error("❌ Falha ao ler a aba de Premissas.")
    st.code(str(e))
    st.stop()

if df.empty:
    st.warning("A base foi carregada, mas não há registros utilizáveis após a limpeza.")
    st.stop()

excel_last_update = get_excel_last_update(EXCEL_FILE)

st.sidebar.header("Filtros")
filtered = df.copy()
filtered = create_safe_multiselect(filtered, "Impress Type", "Impress Type", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "g/m2", "g/m2", numeric_no_decimal=True)
filtered = create_safe_multiselect(filtered, "Supplier", "Supplier", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "Currency", "Currency", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "Lot (ton)", "Lot (ton)", numeric_no_decimal=True)

st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
st.sidebar.markdown(
    f"""
    <div style='font-size:0.9rem;color:#6B7280;'>Cockpit_Papel.xlsx</div>
    <div style='font-size:0.9rem;color:#111827;'><b>Última atualização:</b> {excel_last_update}</div>
    <div style='font-size:0.8rem;color:#6B7280;'>App V1.0 - MLC_2026</div>
    """,
    unsafe_allow_html=True,
)

if filtered.shape[0] == df.shape[0]:
    st.info("Selecione pelo menos um filtro para visualizar os dados.")
    st.stop()

k5, k6, k7, k8, k9 = st.columns(5)
with k5:
    kpi_card("Frete CN", format_br_number(premissas_kpis["Frete CN"], 2) if premissas_kpis["Frete CN"] is not None else "N/A")
with k6:
    kpi_card("Frete EU", format_br_number(premissas_kpis["Frete EU"], 2) if premissas_kpis["Frete EU"] is not None else "N/A")
with k7:
    kpi_card("USD/BRL", format_br_number(premissas_kpis["USD/BRL"], 2) if premissas_kpis["USD/BRL"] is not None else "N/A")
with k8:
    kpi_card("EUR/BRL", format_br_number(premissas_kpis["EUR/BRL"], 2) if premissas_kpis["EUR/BRL"] is not None else "N/A")
with k9:
    kpi_card("CNY/BRL", format_br_number(premissas_kpis["CNY/BRL"], 2) if premissas_kpis["CNY/BRL"] is not None else "N/A")

st.markdown("### Tabela principal")

# Width foi removido da visão principal e dos filtros.
display_cols = [
    "Impress Type",
    "g/m2",
    "Supplier",
    "Current Price",
    "Currency",
    "P.Value (R$/KG)",
    "P.Value (R$/M2)",
    "Última Atualização de Preço",
    "Payment Terms",
    "Lot (ton)",
]
display_cols = [c for c in display_cols if c in filtered.columns]

table_df_raw = filtered[display_cols].copy()
table_df_raw = table_df_raw.rename(columns={"Última Atualização de Preço": "Último Preço"})
table_df_display = table_df_raw.copy()

value_cols = [
    "Current Price",
    "TCO (R$/KG)",
    "TCO (R$/M2)",
    "P.Value (R$/KG)",
    "P.Value (R$/M2)",
]
no_decimal_cols = ["g/m2", "Paper bonus (t)", "Lot (ton)", "Working days", "Payment Terms"]

for col in value_cols:
    if col in table_df_display.columns:
        table_df_display[col] = table_df_display[col].apply(lambda x: format_br_number(x, 2))

for col in no_decimal_cols:
    if col in table_df_display.columns:
        table_df_display[col] = table_df_display[col].apply(lambda x: "" if pd.isna(x) else format_no_decimal(x))

if "Último Preço" in table_df_display.columns:
    table_df_display["Último Preço"] = pd.to_datetime(table_df_display["Último Preço"], errors="coerce").dt.strftime("%d/%m/%Y")

# Exibição simples e estável
st.dataframe(table_df_display, use_container_width=True, hide_index=True)

exp1, exp2, exp3 = st.columns([1.2, 1.2, 6])
with exp1:
    csv_bytes = table_df_display.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="Exportar CSV",
        data=csv_bytes,
        file_name="cockpit_filtrado.csv",
        mime="text/csv",
        use_container_width=True,
    )
with exp2:
    excel_bytes = to_excel_bytes(table_df_display)
    st.download_button(
        label="Exportar Excel",
        data=excel_bytes,
        file_name="cockpit_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("### TCO médio por Supplier")
if {"Supplier", "TCO (R$/KG)"}.issubset(filtered.columns):
    chart_df = filtered.copy()
    chart_df["TCO (R$/KG)"] = pd.to_numeric(chart_df["TCO (R$/KG)"], errors="coerce")
    chart_df = chart_df.dropna(subset=["Supplier", "TCO (R$/KG)"])

    if not chart_df.empty:
        chart_base = (
            chart_df.groupby("Supplier", as_index=False)["TCO (R$/KG)"]
            .mean()
            .sort_values("TCO (R$/KG)", ascending=True)
        )
        chart_base["TCO_label"] = chart_base["TCO (R$/KG)"].apply(lambda x: format_br_number(x, 2))

        fig_bar = px.bar(
            chart_base,
            x="Supplier",
            y="TCO (R$/KG)",
            color="TCO (R$/KG)",
            color_continuous_scale="Blues",
            text="TCO_label",
        )
        fig_bar.update_traces(textposition="outside")
        fig_bar.update_layout(
            showlegend=False,
            height=390,
            margin=dict(t=20, r=20, l=20, b=20),
            xaxis_title="Supplier",
            yaxis_title="TCO (R$/KG)",
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        fig_bar.update_xaxes(showgrid=False)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#E5E7EB")
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Sem dados válidos para exibir o gráfico.")
else:
    st.info("As colunas necessárias para o gráfico não estão disponíveis.")

best_pv_kg = safe_min(filtered["P.Value (R$/KG)"]) if "P.Value (R$/KG)" in filtered.columns else None
best_pv_m2 = safe_min(filtered["P.Value (R$/M2)"]) if "P.Value (R$/M2)" in filtered.columns else None
best_row = safe_best_row(filtered, "P.Value (R$/KG)") if "P.Value (R$/KG)" in filtered.columns else None
best_supplier = best_row["Supplier"] if best_row is not None and "Supplier" in filtered.columns else "N/A"

if "Impress Type" in filtered.columns:
    unique_impress = filtered["Impress Type"].dropna().unique().tolist()
    impress_value = str(unique_impress[0]) if len(unique_impress) == 1 else f"{len(unique_impress)} tipos"
else:
    impress_value = "N/A"

k1, k2, k3, k4 = st.columns(4)
with k1:
    kpi_card("Impress Type", impress_value)
with k2:
    kpi_card("Melhor P.Value (R$/KG)", format_br_number(best_pv_kg, 2) if best_pv_kg is not None else "N/A")
with k3:
    kpi_card("Melhor P.Value (R$/M2)", format_br_number(best_pv_m2, 2) if best_pv_m2 is not None else "N/A")
with k4:
    kpi_card("Melhor Supplier", best_supplier if best_supplier else "N/A")

st.caption(
    "O dashboard considera a aba 'Preços e Condições' como base, "
    "usa apenas 'Current Price' como preço principal e ignora colunas históricas mensais. "
    "Nesta versão, Material Number e Width (mm) foram removidos da camada analítica e, "
    "em caso de duplicidade, o app mantém a linha com a data mais recente em 'Última Atualização de Preço'."
)
