import os
import re
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

# =========================================================
# CONFIG
# =========================================================
ARQUIVO_EXCEL = "CG BOT PY.xlsx"
MINUTOS_POR_PESSOA_DIA = 500.0

st.set_page_config(
    page_title="MES Industrial PRO",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CSS - VISUAL PRO MES
# =========================================================
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: "Segoe UI", sans-serif;
}

.stApp {
    background:
        linear-gradient(135deg, rgba(12,18,28,0.98), rgba(18,28,40,0.98)),
        repeating-linear-gradient(
            45deg,
            rgba(255,255,255,0.02) 0px,
            rgba(255,255,255,0.02) 12px,
            rgba(0,0,0,0.02) 12px,
            rgba(0,0,0,0.02) 24px
        );
    color: #EAF2FF;
}

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 1rem;
    max-width: 1600px;
}

h1, h2, h3 {
    color: #EAF2FF;
    letter-spacing: 0.4px;
}

.kpi-card {
    background: linear-gradient(180deg, rgba(26,35,50,0.90), rgba(17,24,39,0.92));
    border: 1px solid rgba(120,180,255,0.22);
    border-radius: 18px;
    padding: 18px 18px 14px 18px;
    box-shadow: 0 0 18px rgba(0,180,255,0.08);
}

.kpi-title {
    font-size: 0.9rem;
    color: #9EC5FF;
    margin-bottom: 6px;
}

.kpi-value {
    font-size: 1.8rem;
    font-weight: 700;
    color: #FFFFFF;
}

.kpi-sub {
    font-size: 0.85rem;
    color: #AAB7CF;
    margin-top: 6px;
}

.section-card {
    background: linear-gradient(180deg, rgba(22,30,44,0.88), rgba(15,21,32,0.90));
    border: 1px solid rgba(120,180,255,0.18);
    border-radius: 20px;
    padding: 18px;
    box-shadow: 0 0 20px rgba(0,180,255,0.05);
    margin-bottom: 16px;
}

.stDataFrame, .stTable {
    background: rgba(12,18,28,0.25);
    border-radius: 12px;
}

div[data-testid="stMetric"] {
    background: linear-gradient(180deg, rgba(26,35,50,0.90), rgba(17,24,39,0.92));
    border: 1px solid rgba(120,180,255,0.18);
    padding: 10px;
    border-radius: 14px;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(13,19,29,0.98), rgba(10,14,22,0.98));
    border-right: 1px solid rgba(120,180,255,0.12);
}

hr {
    border: none;
    border-top: 1px solid rgba(140,180,255,0.15);
    margin: 0.9rem 0 1.1rem 0;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def _to_float(x):
    try:
        if pd.isna(x):
            return np.nan

        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)

        s = str(x).strip()

        if s == "":
            return np.nan

        s = s.replace(" ", "")

        # Ex.: 1.234,56 -> 1234.56
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")

        return float(s)
    except Exception:
        return np.nan


def _safe_key(texto: str) -> str:
    return re.sub(r"[^0-9a-zA-Z_]+", "_", str(texto))[:80]


def _col_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([np.nan] * len(df), index=df.index)

    obj = df[col]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def _num(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index)
    return _col_series(df, col).apply(_to_float)


def _find_col(df: pd.DataFrame, text: str) -> Optional[str]:
    if df.empty:
        return None
    for c in df.columns:
        if text.lower() in str(c).lower():
            return c
    return None


def _find_first_existing(df: pd.DataFrame, candidates):
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for name in candidates:
        key = str(name).strip().lower()
        if key in normalized:
            return normalized[key]

    for c in df.columns:
        c_lower = str(c).strip().lower()
        for name in candidates:
            if str(name).strip().lower() in c_lower:
                return c
    return None


def _util_color(util: float) -> str:
    if util > 100:
        return "Crítico"
    if util >= 85:
        return "Atenção"
    return "Normal"


def _format_pct(valor: float) -> str:
    return f"{valor:.1f}%"


def _format_num(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _kpi_card(title: str, value: str, sub: str = ""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def _status_text(util: float) -> str:
    if util > 100:
        return "CRÍTICO"
    if util >= 85:
        return "ATENÇÃO"
    return "NORMAL"


# =========================================================
# LOAD EXCEL
# =========================================================
if not os.path.exists(ARQUIVO_EXCEL):
    st.error(f"Arquivo Excel não encontrado: {ARQUIVO_EXCEL}")
    st.stop()

try:
    xls = pd.ExcelFile(ARQUIVO_EXCEL)
except Exception as e:
    st.error(f"Erro ao abrir o Excel: {e}")
    st.stop()

try:
    df0 = pd.read_excel(ARQUIVO_EXCEL, sheet_name=0)
except Exception as e:
    st.error(f"Erro ao ler a planilha principal: {e}")
    st.stop()

try:
    df_ind = pd.read_excel(ARQUIVO_EXCEL, sheet_name="INDIRETOS")
except Exception:
    df_ind = pd.DataFrame()

# =========================================================
# MAPEAMENTO DE COLUNAS
# =========================================================
col_modelo = _find_first_existing(
    df0,
    ["MODELO", "MODEL", "PRODUTO", "ITEM", "DESCRIÇÃO", "DESCRICAO", df0.columns[2] if len(df0.columns) > 2 else ""]
)

col_linha = _find_first_existing(
    df0,
    ["LINHA", "LINE", "SETOR", "POSTO", "RECURSO", df0.columns[5] if len(df0.columns) > 5 else ""]
)

col_tempo = _find_first_existing(
    df0,
    ["TEMPO", "TEMPO PADRÃO", "TEMPO PADRAO", "CICLO", "MIN", "MINUTOS", "TP"]
)

col_cr = _find_first_existing(
    df0,
    ["CR", "CENTRO DE RESULTADO", "CENTRO RESULTADO", "CUSTO"]
)

if col_modelo is None:
    st.error("Não foi possível identificar a coluna de MODELO na planilha principal.")
    st.stop()

if col_linha is None:
    st.error("Não foi possível identificar a coluna de LINHA na planilha principal.")
    st.stop()

# =========================================================
# TÍTULO
# =========================================================
st.markdown("""
<div class="section-card">
    <h1 style="margin-bottom:0.2rem;">MES Industrial PRO</h1>
    <div style="color:#9EC5FF; font-size:1rem;">
        Planejamento de carga, capacidade, gargalos e mão de obra em ambiente industrial
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("## Configurações do Cenário")

    oee = st.slider("OEE do cenário", min_value=0.50, max_value=1.00, value=0.85, step=0.01)
    dias = st.number_input("Dias produtivos", min_value=1, max_value=31, value=22, step=1)
    minutos_por_pessoa = st.number_input(
        "Minutos por pessoa/dia",
        min_value=1.0,
        max_value=1440.0,
        value=float(MINUTOS_POR_PESSOA_DIA),
        step=10.0
    )

    st.markdown("---")
    st.markdown("## Filtros")

    modelos_disponiveis = sorted([x for x in df0[col_modelo].dropna().unique().tolist() if str(x).strip() != ""])
    linhas_disponiveis = sorted([x for x in df0[col_linha].dropna().unique().tolist() if str(x).strip() != ""])

    modelos = st.multiselect("Modelo", modelos_disponiveis)
    linhas = st.multiselect("Linha", linhas_disponiveis)

    if col_cr and col_cr in df0.columns:
        cr_disponiveis = sorted([x for x in df0[col_cr].dropna().unique().tolist() if str(x).strip() != ""])
        filtros_cr = st.multiselect("CR", cr_disponiveis)
    else:
        filtros_cr = []

    st.markdown("---")
    st.caption("A capacidade é ajustada pelo OEE e pela mão de obra informada por linha.")

# =========================================================
# FILTRO BASE
# =========================================================
df = df0.copy()

if modelos:
    df = df[df[col_modelo].isin(modelos)]

if linhas:
    df = df[df[col_linha].isin(linhas)]

if col_cr and filtros_cr:
    df = df[df[col_cr].isin(filtros_cr)]

if df.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")
    st.stop()

# =========================================================
# ENTRADA DE QUANTIDADE POR MODELO
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Plano de Produção por Modelo")

modelos_filtrados = sorted([x for x in df[col_modelo].dropna().unique().tolist() if str(x).strip() != ""])

qty_map = {}
cols_qtd = st.columns(4)

for i, modelo in enumerate(modelos_filtrados):
    with cols_qtd[i % 4]:
        qty_map[modelo] = st.number_input(
            f"{modelo}",
            min_value=0,
            max_value=1000000,
            value=0,
            step=1,
            key=f"qtd_{_safe_key(modelo)}"
        )
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# ENTRADA DE MO POR LINHA
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Mão de Obra por Linha")

linhas_filtradas = sorted([x for x in df[col_linha].dropna().unique().tolist() if str(x).strip() != ""])

mo_map = {}
cols_mo = st.columns(4)

for i, linha in enumerate(linhas_filtradas):
    with cols_mo[i % 4]:
        mo_map[linha] = st.number_input(
            f"{linha}",
            min_value=0,
            max_value=500,
            value=1,
            step=1,
            key=f"mo_{_safe_key(linha)}"
        )
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# CÁLCULOS
# =========================================================
df_calc = df.copy()

if col_tempo:
    df_calc["TEMPO_MIN"] = _num(df_calc, col_tempo).fillna(0.0)
else:
    df_calc["TEMPO_MIN"] = 0.0

df_calc["QTD"] = _col_series(df_calc, col_modelo).map(qty_map).fillna(0).astype(float)
df_calc["MO"] = _col_series(df_calc, col_linha).map(mo_map).fillna(0).astype(float)

df_calc["CARGA_MIN"] = df_calc["TEMPO_MIN"] * df_calc["QTD"]
df_calc["HORAS_CARGA"] = df_calc["CARGA_MIN"] / 60.0

df_calc["CAP_BRUTA_MIN"] = df_calc["MO"] * minutos_por_pessoa * float(dias)
df_calc["CAP_REAL_MIN"] = df_calc["CAP_BRUTA_MIN"] * float(oee)
df_calc["CAP_REAL_H"] = df_calc["CAP_REAL_MIN"] / 60.0

# =========================================================
# AGRUPAMENTO POR LINHA
# =========================================================
agg = (
    df_calc.groupby(col_linha, dropna=False)
    .agg(
        carga_min=("CARGA_MIN", "sum"),
        horas_carga=("HORAS_CARGA", "sum"),
        mo=("MO", "max")
    )
    .reset_index()
)

agg["cap_bruta_min"] = agg["mo"] * minutos_por_pessoa * float(dias)
agg["cap_real_min"] = agg["cap_bruta_min"] * float(oee)
agg["cap_bruta_h"] = agg["cap_bruta_min"] / 60.0
agg["cap_real_h"] = agg["cap_real_min"] / 60.0

agg["utilizacao_pct"] = np.where(
    agg["cap_real_min"] > 0,
    (agg["carga_min"] / agg["cap_real_min"]) * 100.0,
    0.0
)

agg["ociosidade_h"] = np.where(
    agg["cap_real_h"] > agg["horas_carga"],
    agg["cap_real_h"] - agg["horas_carga"],
    0.0
)

agg["deficit_h"] = np.where(
    agg["horas_carga"] > agg["cap_real_h"],
    agg["horas_carga"] - agg["cap_real_h"],
    0.0
)

agg["status"] = agg["utilizacao_pct"].apply(_util_color)
agg["status_texto"] = agg["utilizacao_pct"].apply(_status_text)

agg = agg.sort_values("utilizacao_pct", ascending=False).reset_index(drop=True)

# =========================================================
# GARGALOS
# =========================================================
gargalos = agg[agg["utilizacao_pct"] > 100].copy()

if not gargalos.empty:
    gargalo_principal = str(gargalos.iloc[0][col_linha])
    gargalo_pct = float(gargalos.iloc[0]["utilizacao_pct"])
else:
    gargalo_principal = "Sem gargalo"
    gargalo_pct = 0.0

# =========================================================
# KPIs GERAIS
# =========================================================
total_horas = float(df_calc["HORAS_CARGA"].sum())
total_cap_real_h = float(agg["cap_real_h"].sum())
total_mo = float(agg["mo"].sum())
util_global = (total_horas / total_cap_real_h * 100.0) if total_cap_real_h > 0 else 0.0

col1, col2, col3, col4 = st.columns(4)

with col1:
    _kpi_card("Carga Total", f"{_format_num(total_horas)} h", "Horas necessárias do plano")

with col2:
    _kpi_card("Capacidade Real", f"{_format_num(total_cap_real_h)} h", f"OEE aplicado: {_format_pct(oee * 100)}")

with col3:
    _kpi_card("Utilização Global", _format_pct(util_global), "Carga / capacidade real")

with col4:
    if gargalo_principal == "Sem gargalo":
        _kpi_card("Gargalo Principal", "Sem gargalo", "Nenhuma linha acima de 100%")
    else:
        _kpi_card("Gargalo Principal", gargalo_principal, f"Utilização: {_format_pct(gargalo_pct)}")

st.markdown("<br>", unsafe_allow_html=True)

# =========================================================
# GRÁFICO 1 - UTILIZAÇÃO POR LINHA
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Utilização por Linha")

chart_util = (
    alt.Chart(agg)
    .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
    .encode(
        x=alt.X("utilizacao_pct:Q", title="Utilização (%)"),
        y=alt.Y(f"{col_linha}:N", sort="-x", title="Linha"),
        color=alt.Color(
            "status:N",
            scale=alt.Scale(
                domain=["Normal", "Atenção", "Crítico"],
                range=["#14C38E", "#FFB020", "#FF5A5F"]
            ),
            legend=alt.Legend(title="Status")
        ),
        tooltip=[
            alt.Tooltip(f"{col_linha}:N", title="Linha"),
            alt.Tooltip("horas_carga:Q", title="Carga (h)", format=".2f"),
            alt.Tooltip("cap_real_h:Q", title="Capacidade real (h)", format=".2f"),
            alt.Tooltip("utilizacao_pct:Q", title="Utilização (%)", format=".1f"),
            alt.Tooltip("mo:Q", title="MO", format=".0f"),
        ]
    )
    .properties(height=max(320, 42 * len(agg)))
)

linha_100 = alt.Chart(pd.DataFrame({"x": [100]})).mark_rule(
    color="#66B3FF",
    strokeDash=[6, 4]
).encode(x="x:Q")

st.altair_chart(chart_util + linha_100, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# GRÁFICO 2 - CARGA X CAPACIDADE
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Carga x Capacidade Real por Linha")

agg_melt = agg[[col_linha, "horas_carga", "cap_real_h"]].melt(
    id_vars=[col_linha],
    value_vars=["horas_carga", "cap_real_h"],
    var_name="Tipo",
    value_name="Horas"
)

agg_melt["Tipo"] = agg_melt["Tipo"].replace({
    "horas_carga": "Carga",
    "cap_real_h": "Capacidade Real"
})

chart_cap = (
    alt.Chart(agg_melt)
    .mark_bar(cornerRadius=4)
    .encode(
        x=alt.X("Horas:Q", title="Horas"),
        y=alt.Y(f"{col_linha}:N", sort="-x", title="Linha"),
        color=alt.Color(
            "Tipo:N",
            scale=alt.Scale(domain=["Carga", "Capacidade Real"], range=["#58A6FF", "#14C38E"]),
            legend=alt.Legend(title="")
        ),
        xOffset="Tipo:N",
        tooltip=[
            alt.Tooltip(f"{col_linha}:N", title="Linha"),
            alt.Tooltip("Tipo:N", title="Tipo"),
            alt.Tooltip("Horas:Q", title="Horas", format=".2f"),
        ]
    )
    .properties(height=max(320, 42 * len(agg)))
)

st.altair_chart(chart_cap, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# TABELA EXECUTIVA
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Tabela Executiva por Linha")

tabela_exec = agg.copy()
tabela_exec["utilizacao_pct"] = tabela_exec["utilizacao_pct"].round(1)
tabela_exec["horas_carga"] = tabela_exec["horas_carga"].round(2)
tabela_exec["cap_bruta_h"] = tabela_exec["cap_bruta_h"].round(2)
tabela_exec["cap_real_h"] = tabela_exec["cap_real_h"].round(2)
tabela_exec["ociosidade_h"] = tabela_exec["ociosidade_h"].round(2)
tabela_exec["deficit_h"] = tabela_exec["deficit_h"].round(2)

st.dataframe(
    tabela_exec.rename(columns={
        col_linha: "Linha",
        "horas_carga": "Carga (h)",
        "mo": "MO",
        "cap_bruta_h": "Capacidade Bruta (h)",
        "cap_real_h": "Capacidade Real (h)",
        "utilizacao_pct": "Utilização (%)",
        "ociosidade_h": "Ociosidade (h)",
        "deficit_h": "Déficit (h)",
        "status_texto": "Status"
    }),
    use_container_width=True,
    hide_index=True
)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# DETALHE DO GARGALO
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Diagnóstico do Gargalo")

if gargalos.empty:
    st.success("Nenhuma linha está acima de 100% de utilização no cenário atual.")
else:
    top_gargalos = gargalos[[col_linha, "horas_carga", "cap_real_h", "utilizacao_pct", "deficit_h", "mo"]].copy()
    top_gargalos["utilizacao_pct"] = top_gargalos["utilizacao_pct"].round(1)
    top_gargalos["horas_carga"] = top_gargalos["horas_carga"].round(2)
    top_gargalos["cap_real_h"] = top_gargalos["cap_real_h"].round(2)
    top_gargalos["deficit_h"] = top_gargalos["deficit_h"].round(2)

    st.warning(
        f"Gargalo principal: **{gargalo_principal}**, com utilização de **{_format_pct(gargalo_pct)}**."
    )

    st.dataframe(
        top_gargalos.rename(columns={
            col_linha: "Linha",
            "horas_carga": "Carga (h)",
            "cap_real_h": "Capacidade Real (h)",
            "utilizacao_pct": "Utilização (%)",
            "deficit_h": "Déficit (h)",
            "mo": "MO"
        }),
        use_container_width=True,
        hide_index=True
    )
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# DETALHE DO PLANO POR MODELO
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Plano Detalhado por Modelo")

plano_modelo = (
    df_calc.groupby(col_modelo, dropna=False)
    .agg(
        qtd=("QTD", "max"),
        tempo_min=("TEMPO_MIN", "mean"),
        carga_min=("CARGA_MIN", "sum"),
        horas=("HORAS_CARGA", "sum")
    )
    .reset_index()
    .sort_values("horas", ascending=False)
)

plano_modelo["tempo_min"] = plano_modelo["tempo_min"].round(4)
plano_modelo["carga_min"] = plano_modelo["carga_min"].round(2)
plano_modelo["horas"] = plano_modelo["horas"].round(2)

st.dataframe(
    plano_modelo.rename(columns={
        col_modelo: "Modelo",
        "qtd": "Quantidade",
        "tempo_min": "Tempo Unit. (min)",
        "carga_min": "Carga (min)",
        "horas": "Carga (h)"
    }),
    use_container_width=True,
    hide_index=True
)
st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# INDIRETOS / MOI
# =========================================================
if not df_ind.empty:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Indiretos / MOI")

    df_ind_calc = df_ind.copy()

    col_moi = _find_first_existing(df_ind_calc, ["MOI", "MÃO DE OBRA INDIRETA", "MAO DE OBRA INDIRETA"])
    col_setor_ind = _find_first_existing(df_ind_calc, ["SETOR", "ÁREA", "AREA", "LINHA", "DEPARTAMENTO"])
    col_qtd_ind = _find_first_existing(df_ind_calc, ["QTD", "QUANTIDADE", "TOTAL"])

    if col_moi and col_moi in df_ind_calc.columns:
        df_ind_calc["MOI_TRATADO"] = _num(df_ind_calc, col_moi).fillna(0.0)
    else:
        df_ind_calc["MOI_TRATADO"] = 0.0

    if col_qtd_ind and col_qtd_ind in df_ind_calc.columns:
        df_ind_calc["QTD_TRATADA"] = _num(df_ind_calc, col_qtd_ind).fillna(0.0)
    else:
        df_ind_calc["QTD_TRATADA"] = 0.0

    total_moi = float(df_ind_calc["MOI_TRATADO"].sum())
    total_qtd_ind = float(df_ind_calc["QTD_TRATADA"].sum()) if "QTD_TRATADA" in df_ind_calc.columns else 0.0

    c1, c2 = st.columns(2)
    with c1:
        _kpi_card("MOI Total", _format_num(total_moi), "Soma da planilha de indiretos")
    with c2:
        _kpi_card("Qtd. Indiretos", _format_num(total_qtd_ind), "Total tratado da aba indiretos")

    if col_setor_ind:
        moi_setor = (
            df_ind_calc.groupby(col_setor_ind, dropna=False)
            .agg(moi=("MOI_TRATADO", "sum"))
            .reset_index()
            .sort_values("moi", ascending=False)
        )

        chart_moi = (
            alt.Chart(moi_setor)
            .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
            .encode(
                x=alt.X("moi:Q", title="MOI"),
                y=alt.Y(f"{col_setor_ind}:N", sort="-x", title="Setor"),
                color=alt.value("#58A6FF"),
                tooltip=[
                    alt.Tooltip(f"{col_setor_ind}:N", title="Setor"),
                    alt.Tooltip("moi:Q", title="MOI", format=".2f")
                ]
            )
            .properties(height=max(280, 38 * len(moi_setor)))
        )

        st.altair_chart(chart_moi, use_container_width=True)

    mostrar_cols = df_ind_calc.copy()
    st.dataframe(mostrar_cols, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# =========================================================
# RODAPÉ TÉCNICO
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Premissas do Cálculo")
st.markdown(f"""
- **Carga (min)** = Tempo unitário × Quantidade planejada  
- **Capacidade bruta (min)** = MO × {minutos_por_pessoa:.0f} min/dia × {int(dias)} dias  
- **Capacidade real (min)** = Capacidade bruta × OEE  
- **Utilização (%)** = Carga ÷ Capacidade real × 100  
- Gargalos são linhas com utilização **acima de 100%**
""")
st.markdown('</div>', unsafe_allow_html=True)
