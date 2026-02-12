import os
import re
from typing import Optional, List, Dict

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image


# =========================================================
# CONFIG FIXO (SEM UPLOAD)
# =========================================================
ARQUIVO_EXCEL = "CG BOT PY.xlsx"  # <-- nome exato do seu Excel na mesma pasta do app.py

# =========================================================
# Streamlit config (TEM QUE SER A PRIMEIRA CHAMADA st.*)
# =========================================================
st.set_page_config(
    page_title="Dashboard de Carga Máquina (Simulação de Cenários)",
    layout="wide",
)


# =========================================================
# Helpers
# =========================================================
def _to_float(x):
    """Convert numbers that may come as '4,5' (pt-BR) or as numeric."""
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
    """
    Returns a Series even if df[col_name] is a DataFrame (duplicate column names).
    """
    obj = df[col_name]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def _safe_multiselect(label: str, series_or_df) -> List:
    """
    Accepts Series or DataFrame. If DataFrame is provided, uses first column.
    """
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
    """Find a column by case-insensitive substring match after stripping."""
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
    # industrial-ish: green ok, amber attention, red overload
    if util_pct >= 100:
        return "#D62728"  # red
    if util_pct >= 85:
        return "#FF7F0E"  # orange
    return "#2CA02C"      # green


def _load_logo_image() -> Optional[Image.Image]:
    """
    Procura automaticamente um arquivo de logo na pasta do app.
    Aceita: logo.png / logo.jpg / logo.jpeg / logo (sem extensão).
    """
    candidates = ["logo.png", "logo.jpg", "logo.jpeg", "logo.webp", "logo"]
    for fn in candidates:
        if os.path.exists(fn) and os.path.isfile(fn):
            try:
                return Image.open(fn)
            except Exception:
                pass
    # fallback: qualquer arquivo que comece com "logo"
    for fn in os.listdir("."):
        if fn.lower().startswith("logo") and os.path.isfile(fn):
            try:
                return Image.open(fn)
            except Exception:
                continue
    return None


# =========================================================
# Header + Logo
# =========================================================
logo_img = _load_logo_image()
if logo_img is not None:
    st.image(logo_img, width=180)

st.title("Dashboard de Carga Máquina (Simulação de Cenários)")
st.caption("Simule cenários ajustando OEE, eficiência de MO, capacidade e quantidades por modelo (coluna C).")


# =========================================================
# Carregar EXCEL FIXO
# =========================================================
if not os.path.exists(ARQUIVO_EXCEL):
    st.error(
        f"❌ Não encontrei o arquivo Excel fixo: **{ARQUIVO_EXCEL}**\n\n"
        f"Coloque esse arquivo na mesma pasta do app.py ou altere a variável `ARQUIVO_EXCEL` no topo."
    )
    st.stop()

try:
    df0 = pd.read_excel(ARQUIVO_EXCEL, engine="openpyxl")
except Exception as e:
    st.error(f"❌ Não consegui ler o XLSX: {e}")
    st.stop()

if df0.empty:
    st.warning("O arquivo está vazio.")
    st.stop()


# =========================================================
# Mapear colunas C/F/J/R por posição
# =========================================================
col_C = _col_by_index(df0, 2)    # C
col_F = _col_by_index(df0, 5)    # F
col_J = _col_by_index(df0, 9)    # J
col_R = _col_by_index(df0, 17)   # R

# Colunas importantes por nome
col_qtd_base = _find_col(df0, "QTD BASE")
col_takt_linha = _find_col(df0, "TAKT LINHA")
col_takt_dia = _find_col(df0, "TAKT DIA")


# =========================================================
# Sidebar - Cenário & Capacidade
# =========================================================
with st.sidebar:
    st.header("2) Parâmetros do Cenário")
    oee = st.slider("OEE / Eficiência Máquina", min_value=0.50, max_value=1.00, value=0.85, step=0.01)
    eff_mo = st.slider("Eficiência Mão de Obra", min_value=0.50, max_value=1.00, value=0.90, step=0.01)

    st.divider()
    st.header("3) Capacidade")

    override_cap = st.checkbox("Sobrescrever turnos/dias úteis do Excel", value=True)

    h1 = st.number_input("Horas 1º turno", min_value=0.0, max_value=24.0, value=9.0, step=0.5, disabled=not override_cap)
    h2 = st.number_input("Horas 2º turno", min_value=0.0, max_value=24.0, value=9.0, step=0.5, disabled=not override_cap)
    h3 = st.number_input("Horas 3º turno", min_value=0.0, max_value=24.0, value=0.0, step=0.5, disabled=not override_cap)
    dias_uteis = st.number_input("Dias úteis no período", min_value=1.0, max_value=31.0, value=22.0, step=1.0, disabled=not override_cap)

    st.divider()
    st.header("Filtros (C, F, J, R)")

    # Modelo = coluna C
    if col_C is None:
        st.warning("Não encontrei a coluna C (3ª coluna) no arquivo.")
        sel_modelo_c = []
    else:
        sel_modelo_c = _safe_multiselect(
            f"Modelo (coluna C: {str(col_C).strip()})",
            _col_series(df0, col_C)
        )

    # Outros filtros
    sel_F = _safe_multiselect(f"Coluna F ({str(col_F).strip()})" if col_F else "Coluna F", _col_series(df0, col_F) if col_F else None)
    sel_J = _safe_multiselect(f"Coluna J ({str(col_J).strip()})" if col_J else "Coluna J", _col_series(df0, col_J) if col_J else None)
    sel_R = _safe_multiselect(f"Coluna R ({str(col_R).strip()})" if col_R else "Coluna R", _col_series(df0, col_R) if col_R else None)


# =========================================================
# Quantidade por modelo (campos abertos)
# =========================================================
st.subheader("Quantidade por modelo")
st.caption("Digite a quantidade planejada por modelo (coluna C). Isso escala a carga usando TEMPO INDIVIDUAL (coluna G).")

qty_map: Dict[str, float] = {}
sel_C = sel_modelo_c if isinstance(sel_modelo_c, list) else []

if col_C is None:
    st.info("A coluna C (Modelo) é obrigatória para simular quantidades por modelo.")
elif len(sel_C) == 0:
    st.info("Selecione ao menos um modelo na coluna C para digitar as quantidades.")
else:
    # QTD_BASE por modelo (fallback = 1)
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

    st.markdown("**Digite a quantidade planejada para cada modelo selecionado:**")

    # Campos abertos (um por modelo)
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
# Agrupamento + filtros e preparo do DF
# =========================================================
st.divider()
st.header("Gráficos")

group_choice = st.selectbox(
    "Agrupar barras por",
    options=[c for c in [col_R, col_F, col_J, col_C, ("CR" if "CR" in df0.columns else _find_col(df0, "CR"))] if c is not None],
    index=0 if col_R is not None else 0,
)

filters = {
    col_C: sel_C,
    col_F: sel_F,
    col_J: sel_J,
    col_R: sel_R,
}
df = _apply_filters(df0, filters).copy()


# =========================================================
# CÁLCULO DE CARGA (pedido): TEMPO INDIVIDUAL (coluna G) * QTD digitada
# =========================================================
col_tempo_ind = _find_col(df0, "TEMPO INDIVIDUAL")
if col_tempo_ind is None:
    st.error("Não encontrei a coluna 'TEMPO INDIVIDUAL' (coluna G) no arquivo.")
    st.stop()

if col_C is None:
    st.error("Não encontrei a coluna C (Modelo) no arquivo.")
    st.stop()

df["TEMPO_IND_MIN"] = _num(df, col_tempo_ind).fillna(0.0)
modelo_series = _col_series(df, col_C).astype(str)

# Quantidade padrão: QTD_BASE (se existir) senão 0
if col_qtd_base is not None:
    qtd_padrao = _num(df, col_qtd_base).fillna(0.0)
else:
    qtd_padrao = pd.Series(0.0, index=df.index)

# Quantidade do cenário (digitável por modelo)
if isinstance(qty_map, dict) and len(qty_map) > 0:
    qtd_cenario = modelo_series.map(lambda m: float(qty_map.get(m, np.nan)))
    qtd_cenario = pd.to_numeric(qtd_cenario, errors="coerce").fillna(qtd_padrao)
else:
    qtd_cenario = qtd_padrao

df["QTD_CENARIO"] = pd.to_numeric(qtd_cenario, errors="coerce").fillna(0.0).clip(lower=0.0)
df["CARGA_MIN"] = df["TEMPO_IND_MIN"] * df["QTD_CENARIO"]
df["HORAS_TRABALHADAS"] = df["CARGA_MIN"] / 60.0


# =========================================================
# TAKT (sem filtro): soma (preferir TAKT LINHA; fallback TAKT DIA)
# =========================================================
col_takt = col_takt_linha or col_takt_dia
if col_takt is None:
    df["TAKT_MIN"] = 0.0
else:
    df["TAKT_MIN"] = _num(df, col_takt).fillna(0.0)
df["TAKT_HORAS"] = df["TAKT_MIN"] / 60.0


# =========================================================
# CAPACIDADE (pedido): cada CR * horas do período, CRs APENAS do recorte e SEM duplicidade
# =========================================================
cr_col = "CR" if "CR" in df.columns else (_find_col(df0, "CR"))
if cr_col is None:
    st.error("Não encontrei a coluna 'CR' no arquivo.")
    st.stop()

df["_CR_CLEAN"] = _col_series(df, cr_col).astype(str).str.strip()
df["_CR_CLEAN"] = df["_CR_CLEAN"].replace({"": np.nan, "nan": np.nan, "None": np.nan})

horas_periodo = (h1 + h2 + h3) * float(dias_uteis)

# CRs únicos do recorte filtrado
n_cr_total = df["_CR_CLEAN"].nunique(dropna=True)

cap_horas_programadas = float(n_cr_total) * float(horas_periodo)
cap_horas_efetivas = cap_horas_programadas * float(oee) * float(eff_mo)


# =========================================================
# KPIs
# =========================================================
total_horas = float(df["HORAS_TRABALHADAS"].sum())
util_pct = (total_horas / cap_horas_efetivas * 100.0) if cap_horas_efetivas > 0 else np.nan

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Horas trabalhadas (carga)", f"{total_horas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
kpi2.metric("Capacidade programada (h)", f"{cap_horas_programadas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
kpi3.metric("Capacidade efetiva (h) (OEE×MO)", f"{cap_horas_efetivas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
kpi4.metric("Utilização (%)", f"{util_pct:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".") if not np.isnan(util_pct) else "-")

st.divider()


# =========================================================
# Charts (industrial-ish)
# =========================================================
import altair as alt

if group_choice is None:
    group_choice = col_R or col_F or col_J or col_C

group_name = str(group_choice).strip()

group_series = _col_series(df, group_choice) if group_choice else pd.Series(["(sem grupo)"] * len(df))
tmp = df.copy()
tmp["_GRUPO_"] = group_series.astype(str).fillna("(vazio)")

agg = tmp.groupby("_GRUPO_", dropna=False).agg(
    horas=("HORAS_TRABALHADAS", "sum"),
    takt_h=("TAKT_HORAS", "sum"),
    linhas=("HORAS_TRABALHADAS", "size"),
    n_cr=("_CR_CLEAN", pd.Series.nunique),
).reset_index()

agg = agg.sort_values("horas", ascending=False)

# Capacidade por agrupamento: n_cr * horas_periodo
agg["cap_prog_h"] = agg["n_cr"].astype(float) * float(horas_periodo)
agg["cap_efet_h"] = agg["cap_prog_h"] * float(oee) * float(eff_mo)

agg["util_pct"] = np.where(agg["cap_efet_h"] > 0, agg["horas"] / agg["cap_efet_h"] * 100.0, np.nan)
agg["cor"] = agg["util_pct"].apply(lambda x: _util_color(float(x)) if not np.isnan(x) else "#7F7F7F")

left, right = st.columns([1.2, 1.0], gap="large")

with left:
    st.subheader("Carga (horas trabalhadas) por agrupamento")
    st.caption(f"Agrupado por: {group_name} • Barras coloridas por utilização (vs capacidade efetiva).")

    chart = alt.Chart(agg).mark_bar().encode(
        x=alt.X("horas:Q", title="Horas (carga)"),
        y=alt.Y("_GRUPO_:N", sort="-x", title=""),
        color=alt.Color("cor:N", scale=None, legend=None),
        tooltip=[
            alt.Tooltip("_GRUPO_:N", title="Grupo"),
            alt.Tooltip("horas:Q", title="Horas (carga)", format=",.2f"),
            alt.Tooltip("cap_efet_h:Q", title="Capacidade efetiva (h)", format=",.2f"),
            alt.Tooltip("util_pct:Q", title="Utilização (%)", format=",.1f"),
        ],
    ).properties(height=min(600, 25 * max(8, len(agg))))

    # Linha de referência: capacidade efetiva TOTAL do recorte
    cap_line = alt.Chart(pd.DataFrame({"x": [cap_horas_efetivas]})).mark_rule(strokeDash=[6, 4]).encode(
        x="x:Q"
    )
    st.altair_chart(chart + cap_line, use_container_width=True)

with right:
    st.subheader("Gráfico TAKT (soma)")
    st.caption("Somando TAKT (em horas).")

    chart2 = alt.Chart(agg).mark_bar().encode(
        x=alt.X("takt_h:Q", title="Horas (soma do TAKT)"),
        y=alt.Y("_GRUPO_:N", sort="-x", title=""),
        tooltip=[
            alt.Tooltip("_GRUPO_:N", title="Grupo"),
            alt.Tooltip("takt_h:Q", title="Horas (TAKT somado)", format=",.2f"),
        ],
    ).properties(height=min(600, 25 * max(8, len(agg))))

    st.altair_chart(chart2, use_container_width=True)

st.divider()


# =========================================================
# Detail table
# =========================================================
st.subheader("Detalhes filtrados (carga em horas)")

show_cols = []
for c in [col_C, col_F, col_J, col_R, cr_col, col_qtd_base, col_tempo_ind, col_takt]:
    if c is not None and c in df.columns and c not in show_cols:
        show_cols.append(c)

detail = df[show_cols].copy()
detail["QTD_CENARIO"] = df["QTD_CENARIO"].round(0)
detail["HORAS_TRABALHADAS"] = df["HORAS_TRABALHADAS"].round(3)
detail["TAKT_HORAS"] = df["TAKT_HORAS"].round(3)

st.dataframe(detail, use_container_width=True, height=420)

csv = detail.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Baixar dados filtrados (CSV)",
    data=csv,
    file_name="carga_maquina_filtrada.csv",
    mime="text/csv",
)
