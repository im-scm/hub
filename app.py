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

st.set_page_config(page_title="Cockpit Papel", layout="wide")

EXCEL_FILE = "Cockpit_Papel.xlsm"
SOURCE_SHEET = "Preços e Condições"
PREMISSAS_SHEET = "Premissas"
APP_TITLE = "Cockpit Papel"
APP_VERSION = "v2.3"
TABLE_HEIGHT_PX = 560

st.markdown(
    """
    <style>
        .block-container {padding-top: 1.0rem; padding-bottom: 1.2rem;}
        .mlc-kpi {border: 1px solid #E5E7EB; border-radius: 14px; padding: 14px 16px; background: #FFFFFF; box-shadow: 0 1px 2px rgba(0,0,0,0.04); min-height: 86px;}
        .mlc-kpi-label {font-size: 0.84rem; color: #6B7280; margin-bottom: 6px;}
        .mlc-kpi-value {font-size: 1.00rem; font-weight: 700; color: #111827; line-height: 1.2;}
        .mlc-section-title {margin-top: 0.35rem; margin-bottom: 0.8rem;}
        .mlc-table-shell {border: 1px solid #E5E7EB; border-radius: 14px; background: #FFFFFF; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03);}
        .mlc-table-wrap {max-height: 560px; overflow-y: auto; overflow-x: auto;}
        .mlc-table {width: 100%; border-collapse: separate; border-spacing: 0; table-layout: auto; font-size: 0.84rem; color: #111827;}
        .mlc-table thead th {position: sticky; top: 0; z-index: 2; background: #F9FAFB; color: #6B7280; font-weight: 500; text-align: center; border-bottom: 1px solid #E5E7EB; border-right: 1px solid #F0F2F5; padding: 10px 12px; white-space: nowrap;}
        .mlc-table tbody td {padding: 5px 12px; border-bottom: 1px solid #F3F4F6; border-right: 1px solid #F8FAFC; white-space: nowrap; background: #FFFFFF;}
        .mlc-table tbody tr:hover td {background: #FAFBFF;}
        .mlc-left { text-align: left; }
        .mlc-center { text-align: center; }
        .mlc-right { text-align: right; font-variant-numeric: tabular-nums; }
        .mlc-table thead th:last-child, .mlc-table tbody td:last-child { border-right: none; }
        .mlc-note {color: #6B7280; font-size: 0.82rem; margin-top: 0.45rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <h1 style='font-size: 1.9rem; margin-bottom: 0.2rem; color: #111827;'>
        {APP_TITLE}
    </h1>
    """,
    unsafe_allow_html=True,
)

def normalize_text(value):
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    text = text.replace("\n", " ")
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


def kpi_card(label, value):
    st.markdown(
        f"""
        <div class='mlc-kpi'>
            <div class='mlc-kpi-label'>{escape(str(label))}</div>
            <div class='mlc-kpi-value'>{escape(str(value))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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


def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Cockpit Filtrado")
    output.seek(0)
    return output.getvalue()


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


def deduplicate_for_analysis(df_in):
    df_work = df_in.copy()
    if df_work.empty:
        return df_work
    date_col = "Última Atualização de Preço"
    if date_col in df_work.columns:
        df_work[date_col] = pd.to_datetime(df_work[date_col], errors="coerce", dayfirst=True)
        df_work = df_work.sort_values(by=date_col, ascending=False, na_position="last", kind="stable")
    display_key_cols = [
        "Impress Type", "g/m2", "Supplier", "Current Price", "Currency",
        "P.Value (R$/KG)", "P.Value (R$/M2)", "Payment Terms", "Lot (ton)"
    ]
    key_cols = [c for c in display_key_cols if c in df_work.columns]
    if not key_cols:
        key_cols = [c for c in df_work.columns if c != date_col]
    for col in key_cols:
        if pd.api.types.is_object_dtype(df_work[col]):
            df_work[col] = df_work[col].apply(normalize_text)
    df_work = df_work.drop_duplicates(subset=key_cols, keep="first")
    final_sort = [c for c in ["Impress Type", "g/m2", "Supplier", "Current Price"] if c in df_work.columns]
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
    ordered_cols = [
        "Material Number", "Impress Type", "Width (mm)", "g/m2", "Supplier", "Currency",
        "Current Price", "Paper bonus (t)", "Lot (ton)", "TCO (R$/KG)", "TCO (R$/M2)",
        "Payment Terms", "Working days", "P.Value (R$/KG)", "P.Value (R$/M2)", "Última Atualização de Preço",
    ]
    existing_cols = [c for c in ordered_cols if c in df2.columns]
    df2 = df2[existing_cols].copy()
    numeric_cols = [
        "Width (mm)", "g/m2", "Current Price", "Paper bonus (t)", "Lot (ton)",
        "TCO (R$/KG)", "TCO (R$/M2)", "Working days", "P.Value (R$/KG)", "P.Value (R$/M2)"
    ]
    for col in numeric_cols:
        if col in df2.columns:
            df2[col] = series_to_numeric(df2[col])
    if "Última Atualização de Preço" in df2.columns:
        df2["Última Atualização de Preço"] = pd.to_datetime(df2["Última Atualização de Preço"], errors="coerce", dayfirst=True)
    if "TCO (R$/KG)" not in df2.columns and "Current Price" in df2.columns:
        df2["TCO (R$/KG)"] = df2["Current Price"]
    if "P.Value (R$/KG)" not in df2.columns and "TCO (R$/KG)" in df2.columns:
        df2["P.Value (R$/KG)"] = df2["TCO (R$/KG)"]
    if "TCO (R$/M2)" not in df2.columns and {"TCO (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["TCO (R$/M2)"] = df2["TCO (R$/KG)"] * (df2["g/m2"] / 1000.0)
    if "P.Value (R$/M2)" not in df2.columns and {"P.Value (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["P.Value (R$/M2)"] = df2["P.Value (R$/KG)"] * (df2["g/m2"] / 1000.0)
    essential = [c for c in ["Impress Type", "Supplier"] if c in df2.columns]
    if essential:
        df2 = df2.dropna(subset=essential)
    if "Current Price" in df2.columns:
        df2 = df2[df2["Current Price"].notna()]
    elif "TCO (R$/KG)" in df2.columns:
        df2 = df2[df2["TCO (R$/KG)"].notna()]
    cols_to_drop = [c for c in ["Material Number", "Width (mm)"] if c in df2.columns]
    if cols_to_drop:
        df2 = df2.drop(columns=cols_to_drop)
    df2 = deduplicate_for_analysis(df2)
    return df2


def create_safe_multiselect(df_in, column_name, label, numeric_no_decimal=False):
    if column_name not in df_in.columns:
        return df_in
    source = df_in[column_name].dropna().copy()
    if source.empty:
        return df_in
    opt_df = pd.DataFrame({"original": source})
    opt_df["display"] = opt_df["original"].apply(lambda x: safe_display_string(x, numeric_no_decimal=numeric_no_decimal))
    opt_df = opt_df.drop_duplicates(subset=["display"], keep="first")
    if numeric_no_decimal:
        opt_df["sort_key"] = opt_df["original"].apply(lambda x: parse_number(x) if parse_number(x) is not None else 10**18)
        opt_df = opt_df.sort_values(["sort_key", "display"], kind="stable")
    else:
        opt_df["sort_key"] = opt_df["display"].astype(str)
        opt_df = opt_df.sort_values("sort_key", kind="stable")
    display_options = opt_df["display"].tolist()
    selected_display = st.sidebar.multiselect(label, options=display_options, default=[])
    if not selected_display:
        return df_in
    selected_original = opt_df.loc[opt_df["display"].isin(selected_display), "original"].tolist()
    return df_in[df_in[column_name].isin(selected_original)]


def build_export_table(df_in):
    display_cols = [
        "Impress Type", "g/m2", "Supplier", "Current Price", "Currency",
        "P.Value (R$/KG)", "P.Value (R$/M2)", "Última Atualização de Preço", "Payment Terms", "Lot (ton)"
    ]
    display_cols = [c for c in display_cols if c in df_in.columns]
    table_df = df_in[display_cols].copy().rename(columns={"Última Atualização de Preço": "Último Preço"})

    # ordenação fixa da tabela: menor para maior P.Value (R$/M2)
    if "P.Value (R$/M2)" in table_df.columns:
        sort_series = pd.to_numeric(df_in["P.Value (R$/M2)"], errors="coerce")
        table_df = table_df.assign(_sort_pv_m2=sort_series.values).sort_values(
            by=["_sort_pv_m2", "Impress Type", "g/m2", "Supplier"],
            ascending=[True, True, True, True],
            na_position="last",
            kind="stable",
        ).drop(columns=["_sort_pv_m2"])

    for col in ["Current Price", "P.Value (R$/KG)", "P.Value (R$/M2)"]:
        if col in table_df.columns:
            table_df[col] = table_df[col].apply(lambda x: format_br_number(x, 2))
    for col in ["g/m2", "Payment Terms", "Lot (ton)", "Working days"]:
        if col in table_df.columns:
            table_df[col] = table_df[col].apply(lambda x: "" if pd.isna(x) else format_no_decimal(x))
    if "Último Preço" in table_df.columns:
        table_df["Último Preço"] = pd.to_datetime(table_df["Último Preço"], errors="coerce").dt.strftime("%d/%m/%Y")
    return table_df


def build_html_table(df_html, table_height_px=560):
    if df_html.empty:
        return "<div class='mlc-note'>Sem registros para exibir.</div>"
    numeric_right_cols = {"Current Price", "P.Value (R$/KG)", "P.Value (R$/M2)"}
    left_cols = {"Supplier"}
    html_parts = ["<div class='mlc-table-shell'>", f"<div class='mlc-table-wrap' style='max-height:{int(table_height_px)}px;'>", "<table class='mlc-table'>", "<thead><tr>"]
    for col in df_html.columns:
        cls = "mlc-right" if col in numeric_right_cols else ("mlc-left" if col in left_cols else "mlc-center")
        html_parts.append(f"<th class='{cls}'>{escape(str(col))}</th>")
    html_parts.append("</tr></thead><tbody>")
    for _, row in df_html.iterrows():
        html_parts.append("<tr>")
        for col in df_html.columns:
            value = "" if pd.isna(row[col]) else str(row[col])
            cls = "mlc-right" if col in numeric_right_cols else ("mlc-left" if col in left_cols else "mlc-center")
            html_parts.append(f"<td class='{cls}'>{escape(value)}</td>")
        html_parts.append("</tr>")
    html_parts.extend(["</tbody></table></div></div>"])
    return "".join(html_parts)


@st.cache_data(show_spinner=False)
def load_data_with_audit():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=None)
    header_row = detect_header_row(raw)
    if header_row is None:
        raise ValueError("Não foi possível identificar automaticamente a linha de cabeçalho da aba 'Preços e Condições'.")
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=header_row)
    df_raw.columns = [re.sub(r"\s+", " ", str(c).strip().replace("\n", " ")) for c in df_raw.columns]
    df_raw = df_raw.dropna(how="all")
    original_count = len(df_raw)
    df_final = build_canonical_dataframe(df_raw)
    final_count = len(df_final)
    removed_count = max(original_count - final_count, 0)
    reduction_pct = (removed_count / original_count * 100.0) if original_count else 0.0
    audit = {
        "original_count": original_count,
        "final_count": final_count,
        "removed_count": removed_count,
        "reduction_pct": reduction_pct,
    }
    return df_final, audit


try:
    df, audit_info = load_data_with_audit()
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

with st.sidebar:
    st.markdown("### Assistente de Análise")
    st.divider()
filtered = df.copy()
filtered = create_safe_multiselect(filtered, "Impress Type", "Impress Type")
filtered = create_safe_multiselect(filtered, "g/m2", "g/m2", numeric_no_decimal=True)
filtered = create_safe_multiselect(filtered, "Supplier", "Supplier")
filtered = create_safe_multiselect(filtered, "Currency", "Currency")
filtered = create_safe_multiselect(filtered, "Lot (ton)", "Lot (ton)", numeric_no_decimal=True)
with st.sidebar:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <div style='font-size:0.92rem;color:#111827;'><b>Cockpit_Papel.xlsm</b></div>
        <div style='font-size:0.88rem;color:#4B5563;'>Última atualização: {excel_last_update}</div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(APP_VERSION)

if filtered.shape[0] == df.shape[0]:
    st.info("Selecione pelo menos um filtro para visualizar os dados.")
    st.stop()

# KPIs 1
col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
with col_f1: kpi_card("Frete CN", format_br_number(premissas_kpis["Frete CN"], 2) if premissas_kpis["Frete CN"] is not None else "N/A")
with col_f2: kpi_card("Frete EU", format_br_number(premissas_kpis["Frete EU"], 2) if premissas_kpis["Frete EU"] is not None else "N/A")
with col_f3: kpi_card("USD/BRL", format_br_number(premissas_kpis["USD/BRL"], 2) if premissas_kpis["USD/BRL"] is not None else "N/A")
with col_f4: kpi_card("EUR/BRL", format_br_number(premissas_kpis["EUR/BRL"], 2) if premissas_kpis["EUR/BRL"] is not None else "N/A")
with col_f5: kpi_card("CNY/BRL", format_br_number(premissas_kpis["CNY/BRL"], 2) if premissas_kpis["CNY/BRL"] is not None else "N/A")

st.write("")
table_df_display = build_export_table(filtered)
table_html = build_html_table(table_df_display, table_height_px=TABLE_HEIGHT_PX)
st.markdown(table_html, unsafe_allow_html=True)
st.markdown("<div class='mlc-note'>Tabela ordenada automaticamente do menor para o maior em P.Value (R$/M2). Alinhamento aplicado: números à direita, Supplier à esquerda e demais colunas centralizadas.</div>", unsafe_allow_html=True)

exp1, exp2, _ = st.columns([1.2, 1.2, 6])
with exp1:
    csv_bytes = table_df_display.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(label="Exportar CSV", data=csv_bytes, file_name="cockpit_filtrado.csv", mime="text/csv", width="stretch")
with exp2:
    excel_bytes = to_excel_bytes(table_df_display)
    st.download_button(label="Exportar Excel", data=excel_bytes, file_name="cockpit_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

st.markdown("<div class='mlc-section-title'><h3>TCO médio por Supplier</h3></div>", unsafe_allow_html=True)
if {"Supplier", "TCO (R$/KG)"}.issubset(filtered.columns):
    chart_df = filtered.copy()
    chart_df["TCO (R$/KG)"] = pd.to_numeric(chart_df["TCO (R$/KG)"], errors="coerce")
    chart_df = chart_df.dropna(subset=["Supplier", "TCO (R$/KG)"])
    if not chart_df.empty:
        chart_base = chart_df.groupby("Supplier", as_index=False)["TCO (R$/KG)"].mean().sort_values("TCO (R$/KG)", ascending=True)
        chart_base["TCO_label"] = chart_base["TCO (R$/KG)"].apply(lambda x: format_br_number(x, 2))
        fig_bar = px.bar(chart_base, x="Supplier", y="TCO (R$/KG)", color="TCO (R$/KG)", color_continuous_scale="Blues", text="TCO_label")
        fig_bar.update_traces(textposition="outside")
        fig_bar.update_layout(showlegend=False, height=420, margin=dict(t=10, r=20, l=20, b=20), xaxis_title="Supplier", yaxis_title="TCO (R$/KG)", plot_bgcolor="white", paper_bgcolor="white")
        fig_bar.update_xaxes(showgrid=False)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#E5E7EB")
        st.plotly_chart(fig_bar, width="stretch")
        st.caption("Fonte: aba 'Preços e Condições' do arquivo Cockpit_Papel.xlsm, após remoção de linhas obsoletas e limpeza analítica.")
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
latest_price_dt = pd.to_datetime(filtered["Última Atualização de Preço"], errors="coerce").max() if "Última Atualização de Preço" in filtered.columns else None
latest_price_str = latest_price_dt.strftime("%d/%m/%Y") if pd.notna(latest_price_dt) else "N/A"
record_count = len(filtered)

k1, k2, k3, k4, k5 = st.columns(5)
with k1: kpi_card("Impress Type", impress_value)
with k2: kpi_card("Melhor P.Value (R$/KG)", format_br_number(best_pv_kg, 2) if best_pv_kg is not None else "N/A")
with k3: kpi_card("Melhor P.Value (R$/M2)", format_br_number(best_pv_m2, 2) if best_pv_m2 is not None else "N/A")
with k4: kpi_card("Melhor Supplier", best_supplier if best_supplier else "N/A")
with k5: kpi_card("Último preço considerado", latest_price_str)

st.caption(f"Registros exibidos após filtros: {record_count}. Critério de remoção de obsolescência: se duas linhas têm a mesma visão exibida na tabela principal, o app mantém apenas a mais recente em 'Última Atualização de Preço'.")
