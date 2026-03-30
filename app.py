import os
import re
from typing import Optional, List, Dict

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
import altair as alt


# =========================================================
# CONFIG FIXO
# =========================================================
ARQUIVO_EXCEL = "CG BOT PY.xlsx"   # nome exato do Excel na mesma pasta
MINUTOS_POR_PESSOA_DIA = 500.0     # regra solicitada


# =========================================================
# STREAMLIT CONFIG
# =========================================================
st.set_page_config(
    page_title="Dashboard de Carga Máquina e Mão de Obra",
    layout="wide",
)


# =========================================================
# HELPERS
# =========================================================
def _to_float(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return np.nan
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan


def _col_by_index(df: pd.DataFrame, idx0: int) -> Optional[str]:
    if df is None or df.empty:
        return None
    if idx0 < 0 or idx0 >= df.shape[1]:
        return None
    return df.columns[idx0]


def _col_series(df: pd.DataFrame, col_name: str) -> pd.Series:
    obj = df[col_name]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def _safe_multiselect(label: str, series_or_df) -> List:
    x = series_or_df
    if isinstance(x, pd.DataFrame):
        x = x.iloc[:, 0]
    if x is None:
        return []
    x = x.dropna()
    vals = pd.unique(x)
    try:
        vals = sorted(vals)
    except Exception:
        vals = sorted([str(v) for v in vals])
    return st.multiselect(label, vals)


def _num(df: pd.DataFrame, col_name: str) -> pd.Series:
    s = _col_series(df, col_name)
    return s.apply(_to_float)


def _find_col(df: pd.DataFrame, contains: str) -> Optional[str]:
    contains_norm = contains.strip().lower()
    candidates = []
    for c in df.columns:
        if not isinstance(c, str):
            continue
        c_norm = c.strip().lower()
        if contains_norm in c_norm:
            candidates.append(c)
    if not candidates:
        return None
    candidates.sort(key=lambda x: len(str(x)))
    return candidates[0]


def _apply_filters(df: pd.DataFrame, filters: Dict[str, List]) -> pd.DataFrame:
    out = df.copy()
    for col, selected in filters.items():
        if col is None or selected is None or len(selected) == 0:
            continue
        s = _col_series(out, col)
        out = out[s.isin(selected)]
    return out


def _util_color(util_pct: float) -> str:
    if util_pct >= 100:
        return "#D62728"
    if util_pct >= 85:
        return "#FF7F0E"
    return "#2CA02C"


def _load_logo_image() -> Optional[Image.Image]:
    candidates = ["logo.png", "logo.jpg", "logo.jpeg", "logo.webp", "logo"]
    for fn in candidates:
        if os.path.exists(fn) and os.path.isfile(fn):
            try:
                return Image.open(fn)
            except Exception:
                pass
    for fn in os.listdir("."):
        if fn.lower().startswith("logo") and os.path.isfile(fn):
            try:
                return Image.open(fn)
            except Exception:
                continue
    return None


def _fmt_br(x, casas=2):
    if pd.isna(x):
        return "-"
    return f"{x:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


# =========================================================
# HEADER
# =========================================================
logo_img = _load_logo_image()
if logo_img is not None:
    st.image(logo_img, width=180)

st.title("Dashboard de Carga Máquina e Simulação de Mão de Obra")
st.caption("Base fixa no Excel da pasta. Simule carga máquina e mão de obra em abas separadas.")


# =========================================================
# CARREGAR EXCEL FIXO
# =========================================================
if not os.path.exists(ARQUIVO_EXCEL):
    st.error(f"Arquivo não encontrado: {ARQUIVO_EXCEL}")
    st.stop()

try:
    df0 = pd.read_excel(ARQUIVO_EXCEL, engine="openpyxl")
except Exception as e:
    st.error(f"Não consegui ler o arquivo Excel: {e}")
    st.stop()

if df0.empty:
    st.warning("O arquivo está vazio.")
    st.stop()


# =========================================================
# MAPEAR COLUNAS
# =========================================================
col_C = _col_by_index(df0, 2)    # C = Modelo
col_F = _col_by_index(df0, 5)    # F = Descrição CR
col_J = _col_by_index(df0, 9)    # J
col_R = _col_by_index(df0, 17)   # R

col_qtd_base = _find_col(df0, "QTD BASE")
col_tempo_ind = _find_col(df0, "TEMPO INDIVIDUAL")
col_takt_linha = _find_col(df0, "TAKT LINHA")
col_takt_dia = _find_col(df0, "TAKT DIA")
cr_col = "CR" if "CR" in df0.columns else _find_col(df0, "CR")

if col_tempo_ind is None:
    st.error("Não encontrei a coluna de TEMPO INDIVIDUAL (coluna G).")
    st.stop()

if col_C is None:
    st.error("Não encontrei a coluna C (Modelo).")
    st.stop()

if col_F is None:
    st.error("Não encontrei a coluna F (Descrição CR).")
    st.stop()


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("1) Cenário")

    oee = st.slider("OEE / Eficiência Máquina", min_value=0.50, max_value=1.00, value=0.85, step=0.01)
    eff_mo = st.slider("Eficiência Mão de Obra", min_value=0.50, max_value=1.00, value=0.90, step=0.01)

    st.divider()
    st.header("2) Capacidade")

    h1 = st.number_input("Horas 1º turno", min_value=0.0, max_value=24.0, value=9.0, step=0.5)
    h2 = st.number_input("Horas 2º turno", min_value=0.0, max_value=24.0, value=9.0, step=0.5)
    h3 = st.number_input("Horas 3º turno", min_value=0.0, max_value=24.0, value=0.0, step=0.5)
    dias_uteis = st.number_input("Dias úteis no período", min_value=1.0, max_value=31.0, value=22.0, step=1.0)

    st.divider()
    st.header("3) Filtros")

    sel_modelo_c = _safe_multiselect(
        f"Modelo (coluna C: {str(col_C).strip()})",
        _col_series(df0, col_C)
    )

    sel_F = _safe_multiselect(
        f"Descrição CR (coluna F: {str(col_F).strip()})",
        _col_series(df0, col_F)
    )

    sel_J = _safe_multiselect(
        f"Coluna J ({str(col_J).strip()})" if col_J else "Coluna J",
        _col_series(df0, col_J) if col_J else None
    )

    sel_R = _safe_multiselect(
        f"Coluna R ({str(col_R).strip()})" if col_R else "Coluna R",
        _col_series(df0, col_R) if col_R else None
    )


# =========================================================
# QUANTIDADE POR MODELO
# =========================================================
st.subheader("Quantidade por modelo")
st.caption("Digite a quantidade planejada por modelo. A base de cálculo usa o tempo individual da coluna G.")

qty_map: Dict[str, float] = {}
sel_C = sel_modelo_c if isinstance(sel_modelo_c, list) else []

if len(sel_C) == 0:
    st.info("Selecione ao menos um modelo para informar a quantidade planejada.")
else:
    if col_qtd_base:
        base_series = pd.to_numeric(_col_series(df0, col_qtd_base), errors="coerce")
        base_by_model = (
            pd.DataFrame({
                "MODELO": _col_series(df0, col_C).astype(str),
                "QTD_BASE": base_series,
            })
            .groupby("MODELO")["QTD_BASE"]
            .first()
        )
    else:
        base_by_model = pd.Series(dtype=float)

    cols_ui = st.columns(2) if len(sel_C) > 8 else [None]
    use_two_cols = len(cols_ui) == 2

    for i, m in enumerate(sel_C):
        m_str = str(m)
        base_val = float(base_by_model.get(m_str, np.nan)) if hasattr(base_by_model, "get") else np.nan
        if not np.isfinite(base_val) or base_val <= 0:
            base_val = 1.0

        key = "qtd_modelo__" + re.sub(r"[^0-9a-zA-Z_]+", "_", m_str)[:80]
        target = cols_ui[i % 2] if use_two_cols else st

        qtd_plan = target.number_input(
            label=f"{m_str}",
            min_value=0,
            value=int(round(base_val)),
            step=1,
            key=key,
        )
        qty_map[m_str] = float(qtd_plan)


# =========================================================
# FILTRAR BASE
# =========================================================
filters = {
    col_C: sel_C,
    col_F: sel_F,
    col_J: sel_J,
    col_R: sel_R,
}
df = _apply_filters(df0, filters).copy()

if df.empty:
    st.warning("Nenhum registro encontrado com os filtros selecionados.")
    st.stop()


# =========================================================
# CÁLCULOS BASE
# =========================================================
df["TEMPO_IND_MIN"] = _num(df, col_tempo_ind).fillna(0.0)
modelo_series = _col_series(df, col_C).astype(str)

if col_qtd_base is not None:
    qtd_padrao = _num(df, col_qtd_base).fillna(0.0)
else:
    qtd_padrao = pd.Series(0.0, index=df.index)

if isinstance(qty_map, dict) and len(qty_map) > 0:
    qtd_cenario = modelo_series.map(lambda m: float(qty_map.get(m, np.nan)))
    qtd_cenario = pd.to_numeric(qtd_cenario, errors="coerce").fillna(qtd_padrao)
else:
    qtd_cenario = qtd_padrao

df["QTD_CENARIO"] = pd.to_numeric(qtd_cenario, errors="coerce").fillna(0.0).clip(lower=0.0)
df["CARGA_MIN"] = df["TEMPO_IND_MIN"] * df["QTD_CENARIO"]
df["HORAS_TRABALHADAS"] = df["CARGA_MIN"] / 60.0

col_takt = col_takt_linha or col_takt_dia
if col_takt is None:
    df["TAKT_MIN"] = 0.0
else:
    df["TAKT_MIN"] = _num(df, col_takt).fillna(0.0)
df["TAKT_HORAS"] = df["TAKT_MIN"] / 60.0

horas_periodo = (h1 + h2 + h3) * float(dias_uteis)

if cr_col is not None:
    df["_CR_CLEAN"] = _col_series(df, cr_col).astype(str).str.strip()
    df["_CR_CLEAN"] = df["_CR_CLEAN"].replace({"": np.nan, "nan": np.nan, "None": np.nan})
    n_cr_total = df["_CR_CLEAN"].nunique(dropna=True)
else:
    df["_CR_CLEAN"] = np.nan
    n_cr_total = 0

cap_horas_programadas = float(n_cr_total) * float(horas_periodo)
cap_horas_efetivas = cap_horas_programadas * float(oee) * float(eff_mo)


# =========================================================
# ABAS
# =========================================================
aba1, aba2 = st.tabs(["Carga Máquina", "Mão de Obra"])


# =========================================================
# ABA 1 - CARGA MÁQUINA
# =========================================================
with aba1:
    total_horas = float(df["HORAS_TRABALHADAS"].sum())
    util_pct = (total_horas / cap_horas_efetivas * 100.0) if cap_horas_efetivas > 0 else np.nan

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Horas trabalhadas (carga)", _fmt_br(total_horas))
    k2.metric("Capacidade programada (h)", _fmt_br(cap_horas_programadas))
    k3.metric("Capacidade efetiva (h)", _fmt_br(cap_horas_efetivas))
    k4.metric("Utilização (%)", f"{_fmt_br(util_pct, 1)}%" if not np.isnan(util_pct) else "-")

    st.divider()

    group_choice = st.selectbox(
        "Agrupar barras por",
        options=[c for c in [col_R, col_F, col_J, col_C, cr_col] if c is not None],
        index=0,
        key="group_choice_maquina"
    )

    tmp = df.copy()
    tmp["_GRUPO_"] = _col_series(df, group_choice).astype(str).fillna("(vazio)")

    agg = tmp.groupby("_GRUPO_", dropna=False).agg(
        horas=("HORAS_TRABALHADAS", "sum"),
        takt_h=("TAKT_HORAS", "sum"),
        n_cr=("_CR_CLEAN", pd.Series.nunique),
    ).reset_index()

    agg = agg.sort_values("horas", ascending=False)
    agg["cap_prog_h"] = agg["n_cr"].astype(float) * float(horas_periodo)
    agg["cap_efet_h"] = agg["cap_prog_h"] * float(oee) * float(eff_mo)
    agg["util_pct"] = np.where(agg["cap_efet_h"] > 0, agg["horas"] / agg["cap_efet_h"] * 100.0, np.nan)
    agg["cor"] = agg["util_pct"].apply(lambda x: _util_color(float(x)) if not np.isnan(x) else "#7F7F7F")

    c1, c2 = st.columns([1.2, 1.0], gap="large")

    with c1:
        st.subheader("Carga por agrupamento")
        chart = alt.Chart(agg).mark_bar().encode(
            x=alt.X("horas:Q", title="Horas (carga)"),
            y=alt.Y("_GRUPO_:N", sort="-x", title=""),
            color=alt.Color("cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("_GRUPO_:N", title="Grupo"),
                alt.Tooltip("horas:Q", title="Horas", format=",.2f"),
                alt.Tooltip("cap_efet_h:Q", title="Cap. efetiva", format=",.2f"),
                alt.Tooltip("util_pct:Q", title="Utilização %", format=",.1f"),
            ],
        ).properties(height=min(600, 25 * max(8, len(agg))))

        cap_line = alt.Chart(pd.DataFrame({"x": [cap_horas_efetivas]})).mark_rule(strokeDash=[6, 4]).encode(x="x:Q")
        st.altair_chart(chart + cap_line, use_container_width=True)

    with c2:
        st.subheader("TAKT (soma)")
        chart2 = alt.Chart(agg).mark_bar().encode(
            x=alt.X("takt_h:Q", title="Horas TAKT"),
            y=alt.Y("_GRUPO_:N", sort="-x", title=""),
            tooltip=[
                alt.Tooltip("_GRUPO_:N", title="Grupo"),
                alt.Tooltip("takt_h:Q", title="TAKT somado", format=",.2f"),
            ],
        ).properties(height=min(600, 25 * max(8, len(agg))))
        st.altair_chart(chart2, use_container_width=True)

    st.divider()

    st.subheader("Detalhes filtrados")
    show_cols = []
    for c in [col_C, col_F, col_J, col_R, cr_col, col_qtd_base, col_tempo_ind, col_takt]:
        if c is not None and c in df.columns and c not in show_cols:
            show_cols.append(c)

    detail = df[show_cols].copy()
    detail["QTD_CENARIO"] = df["QTD_CENARIO"].round(0)
    detail["HORAS_TRABALHADAS"] = df["HORAS_TRABALHADAS"].round(3)
    detail["TAKT_HORAS"] = df["TAKT_HORAS"].round(3)

    st.dataframe(detail, use_container_width=True, height=420)


# =========================================================
# ABA 2 - MÃO DE OBRA
# =========================================================
with aba2:
    st.subheader("Simulação de Cenário - Mão de Obra")
    st.caption("Base: tempo individual da coluna G. Regra: 1 pessoa = 500 minutos por dia. Filtro principal: Descrição CR da coluna F.")

    dias_mo = st.number_input(
        "Dias para cálculo da mão de obra",
        min_value=1.0,
        max_value=31.0,
        value=float(dias_uteis),
        step=1.0,
        key="dias_mo"
    )

    minutos_disponiveis_por_pessoa = MINUTOS_POR_PESSOA_DIA * float(dias_mo)

    df_mo = df.copy()
    df_mo["_DESC_CR_"] = _col_series(df_mo, col_F).astype(str).str.strip()

    agg_mo = df_mo.groupby("_DESC_CR_", dropna=False).agg(
        minutos_totais=("CARGA_MIN", "sum"),
        modelos=(col_C, pd.Series.nunique),
        linhas=("CARGA_MIN", "size"),
    ).reset_index()

    agg_mo = agg_mo.sort_values("minutos_totais", ascending=False)
    agg_mo["pessoas_necessarias"] = np.where(
        minutos_disponiveis_por_pessoa > 0,
        agg_mo["minutos_totais"] / minutos_disponiveis_por_pessoa,
        np.nan
    )
    agg_mo["pessoas_arredondadas"] = np.ceil(agg_mo["pessoas_necessarias"].fillna(0)).astype(int)

    total_min_mo = float(agg_mo["minutos_totais"].sum())
    total_pessoas = float(agg_mo["pessoas_necessarias"].sum())
    total_pessoas_arr = int(np.ceil(total_pessoas)) if np.isfinite(total_pessoas) else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Minutos totais", _fmt_br(total_min_mo))
    m2.metric("Minutos por pessoa no período", _fmt_br(minutos_disponiveis_por_pessoa))
    m3.metric("Pessoas necessárias", _fmt_br(total_pessoas, 2))
    m4.metric("Pessoas arredondadas", f"{total_pessoas_arr}")

    st.divider()

    graf_mo = alt.Chart(agg_mo).mark_bar().encode(
        x=alt.X("pessoas_necessarias:Q", title="Pessoas necessárias"),
        y=alt.Y("_DESC_CR_:N", sort="-x", title="Descrição CR"),
        tooltip=[
            alt.Tooltip("_DESC_CR_:N", title="Descrição CR"),
            alt.Tooltip("minutos_totais:Q", title="Minutos totais", format=",.2f"),
            alt.Tooltip("pessoas_necessarias:Q", title="Pessoas necessárias", format=",.2f"),
            alt.Tooltip("pessoas_arredondadas:Q", title="Pessoas arredondadas"),
        ],
    ).properties(height=min(700, 28 * max(8, len(agg_mo))))

    st.altair_chart(graf_mo, use_container_width=True)

    st.divider()

    st.subheader("Tabela de mão de obra por Descrição CR")
    tabela_mo = agg_mo.rename(columns={
        "_DESC_CR_": "DESCRIÇÃO CR",
        "minutos_totais": "MINUTOS_TOTAIS",
        "modelos": "MODELOS",
        "linhas": "LINHAS",
        "pessoas_necessarias": "PESSOAS_NECESSARIAS",
        "pessoas_arredondadas": "PESSOAS_ARREDONDADAS",
    }).copy()

    st.dataframe(tabela_mo, use_container_width=True, height=450)

    csv_mo = tabela_mo.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar mão de obra (CSV)",
        data=csv_mo,
        file_name="mao_de_obra_por_descricao_cr.csv",
        mime="text/csv",
        key="download_mo"
    )
