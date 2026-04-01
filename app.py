import os
import re
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt
from PIL import Image
from typing import Optional, List, Dict

# =========================================================
# CONFIG
# =========================================================
ARQUIVO_EXCEL = "CG BOT PY.xlsx"
MINUTOS_POR_PESSOA_DIA = 500.0

st.set_page_config(
    page_title="MES Industrial",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# FUNÇÕES BASE (CORRIGIDO ERRO)
# =========================================================
def _to_float(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()
    if s == "":
        return np.nan

    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return np.nan


def _safe_key(texto: str) -> str:
    return re.sub(r"[^0-9a-zA-Z_]+", "_", str(texto))[:80]


def _col_series(df, col):
    obj = df[col]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def _num(df, col):
    return _col_series(df, col).apply(_to_float)


def _find_col(df, text):
    for c in df.columns:
        if text.lower() in str(c).lower():
            return c
    return None


def _apply_filters(df, filters):
    out = df.copy()
    for col, vals in filters.items():
        if col and vals:
            out = out[_col_series(out, col).isin(vals)]
    return out


def _util_color(x):
    if x > 100:
        return "#FF5A5F"
    if x >= 85:
        return "#FFB020"
    return "#14C38E"


# =========================================================
# LOAD EXCEL
# =========================================================
if not os.path.exists(ARQUIVO_EXCEL):
    st.error("Arquivo não encontrado.")
    st.stop()

df0 = pd.read_excel(ARQUIVO_EXCEL)

# indiretos
try:
    df_ind = pd.read_excel(ARQUIVO_EXCEL, sheet_name="INDIRETOS")
except:
    df_ind = pd.DataFrame()

# =========================================================
# COLUNAS
# =========================================================
col_C = df0.columns[2]
col_F = df0.columns[5]
col_tempo = _find_col(df0, "TEMPO")
cr_col = _find_col(df0, "CR")

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("Cenário")

    oee = st.slider("OEE", 0.5, 1.0, 0.85)
    dias = st.number_input("Dias", 1, 31, 22)

    st.header("Filtros")
    modelos = st.multiselect("Modelo", df0[col_C].dropna().unique())
    linhas = st.multiselect("Linha", df0[col_F].dropna().unique())

# =========================================================
# FILTRO
# =========================================================
df = df0.copy()

if modelos:
    df = df[df[col_C].isin(modelos)]

if linhas:
    df = df[df[col_F].isin(linhas)]

# =========================================================
# QUANTIDADE POR MODELO
# =========================================================
st.subheader("Quantidade por modelo")

qty_map = {}
for m in df[col_C].dropna().unique():
    qty_map[m] = st.number_input(f"{m}", 0, 100000, 0)

# =========================================================
# MO POR LINHA (CORREÇÃO DO GARGALO)
# =========================================================
st.subheader("Mão de obra por linha")

mo_map = {}
for l in df[col_F].dropna().unique():
    mo_map[l] = st.number_input(f"{l}", 0, 100, 1)

# =========================================================
# CALCULOS
# =========================================================
df["TEMPO"] = _num(df, col_tempo).fillna(0)

df["QTD"] = df[col_C].map(qty_map).fillna(0)

df["CARGA_MIN"] = df["TEMPO"] * df["QTD"]
df["HORAS"] = df["CARGA_MIN"] / 60

df["MO"] = df[col_F].map(mo_map).fillna(0)

df["CAP_MO"] = df["MO"] * MINUTOS_POR_PESSOA_DIA * dias

# =========================================================
# AGRUPAMENTO
# =========================================================
agg = df.groupby(col_F).agg(
    carga_min=("CARGA_MIN", "sum"),
    horas=("HORAS", "sum"),
    mo=("MO", "max")
).reset_index()

agg["cap_mo"] = agg["mo"] * MINUTOS_POR_PESSOA_DIA * dias

agg["util_mo"] = np.where(
    agg["cap_mo"] > 0,
    agg["carga_min"] / agg["cap_mo"] * 100,
    0
)

agg["util"] = agg["util_mo"]  # gargalo real = MO agora

# =========================================================
# GARGALO REAL
# =========================================================
gargalos = agg[agg["util"] > 100].sort_values("util", ascending=False)

gargalo = gargalos.iloc[0][col_F] if not gargalos.empty else "SEM GARGALO"

# =========================================================
# KPIs
# =========================================================
st.metric("Gargalo", gargalo)
st.metric("Carga Total (h)", round(df["HORAS"].sum(), 2))

# =========================================================
# GRÁFICO
# =========================================================
agg["cor"] = agg["util"].apply(_util_color)

chart = alt.Chart(agg).mark_bar().encode(
    x="horas",
    y=col_F,
    color=alt.Color("cor:N", scale=None),
    tooltip=["horas", "util"]
)

st.altair_chart(chart, use_container_width=True)

# =========================================================
# TABELA
# =========================================================
st.dataframe(agg)

# =========================================================
# INDIRETOS
# =========================================================
if not df_ind.empty:
    st.subheader("Indiretos (MOI)")
    st.dataframe(df_ind)
