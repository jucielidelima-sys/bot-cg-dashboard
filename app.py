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
ARQUIVO_EXCEL = "CG BOT PY.xlsx"
MINUTOS_POR_PESSOA_DIA = 500.0


# =========================================================
# CONFIG STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Dashboard de Carga Máquina e Mão de Obra",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# CSS FINAL (SEM LINHA BRANCA + INDUSTRIAL)
# =========================================================
st.markdown("""
<style>
/* REMOVE ESPAÇO BRANCO SUPERIOR */
html, body, [class*="css"]  {
    margin: 0 !important;
    padding: 0 !important;
}

/* REMOVE HEADER DO STREAMLIT */
header[data-testid="stHeader"] {
    background: transparent !important;
    height: 0px !important;
}
header[data-testid="stHeader"] > div {
    height: 0px !important;
}

/* REMOVE ESPAÇO DO CONTAINER */
.block-container {
    padding-top: 0rem !important;
    padding-bottom: 1rem;
    max-width: 96%;
}

/* REMOVE GAP EXTRA */
div[data-testid="stAppViewContainer"] > .main {
    padding-top: 0rem !important;
}
section.main > div {
    padding-top: 0rem !important;
}

/* BACKGROUND INDUSTRIAL */
.stApp {
    background:
        radial-gradient(circle at top left, rgba(45,156,255,0.10), transparent 26%),
        radial-gradient(circle at top right, rgba(139,92,246,0.10), transparent 24%),
        linear-gradient(180deg, #0a0d12 0%, #0f141d 50%, #131925 100%);
    color: #E8EDF7;
}

/* HEADER METÁLICO */
.metal-header {
    border-radius: 20px;
    padding: 18px 24px;
    margin-top: 0rem;
    margin-bottom: 18px;
    background:
        linear-gradient(135deg, #616975, #2e3540, #8b949f, #232a34);
    box-shadow:
        0 8px 25px rgba(0,0,0,0.4),
        inset 0 1px 0 rgba(255,255,255,0.2);
}

.metal-title {
    font-size: 1.9rem;
    font-weight: 900;
    color: white;
}

.metal-subtitle {
    font-size: 0.9rem;
    color: #d6dde8;
}

/* CARDS */
.tesla-card {
    border-radius: 16px;
    padding: 16px;
    margin-bottom: 12px;
    background: rgba(255,255,255,0.05);
    box-shadow: 0 0 15px rgba(45,156,255,0.1);
}

.card-blue { border-left: 5px solid #2D9CFF; }
.card-green { border-left: 5px solid #14C38E; }
.card-orange { border-left: 5px solid #FFB020; }
.card-red { border-left: 5px solid #FF5A5F; }
.card-purple { border-left: 5px solid #8B5CF6; }

.card-title {
    font-size: 0.8rem;
    color: #9fb0c7;
    font-weight: bold;
}

.card-value {
    font-size: 1.8rem;
    font-weight: 900;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================
def _to_float(x):
    if pd.isna(x):
        return np.nan
    try:
        return float(str(x).replace(",", "."))
    except:
        return np.nan


def _col_by_index(df, idx):
    try:
        return df.columns[idx]
    except:
        return None


def _fmt(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# =========================================================
# HEADER
# =========================================================
col_logo, col_head = st.columns([0.15, 0.85])

with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=120)

with col_head:
    st.markdown("""
    <div class="metal-header">
        <div class="metal-title">Dashboard de Carga Máquina e Mão de Obra</div>
        <div class="metal-subtitle">Sala de controle industrial • Simulação de cenários</div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# LOAD DATA
# =========================================================
df = pd.read_excel(ARQUIVO_EXCEL)

col_C = _col_by_index(df, 2)
col_F = _col_by_index(df, 5)
col_G = _col_by_index(df, 6)

df["TEMPO"] = df[col_G].apply(_to_float)

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("Filtros")

    modelos = st.multiselect("Modelo", df[col_C].dropna().unique())
    cr = st.multiselect("CR", df[col_F].dropna().unique())

# =========================================================
# FILTRO
# =========================================================
if modelos:
    df = df[df[col_C].isin(modelos)]

if cr:
    df = df[df[col_F].isin(cr)]

# =========================================================
# QUANTIDADE
# =========================================================
st.subheader("Quantidade por modelo")

qty_map = {}
for m in modelos:
    qty_map[m] = st.number_input(f"{m}", 0, 100000, 100)

df["QTD"] = df[col_C].map(qty_map).fillna(0)
df["MIN"] = df["TEMPO"] * df["QTD"]
df["HORAS"] = df["MIN"] / 60

# =========================================================
# KPIS
# =========================================================
total_horas = df["HORAS"].sum()

c1, c2 = st.columns(2)
c1.metric("Horas totais", _fmt(total_horas))
c2.metric("Registros", len(df))


# =========================================================
# GRÁFICO NEON
# =========================================================
st.subheader("Carga por CR")

agg = df.groupby(col_F)["HORAS"].sum().reset_index()

base = alt.Chart(agg).encode(
    x="HORAS",
    y=alt.Y(col_F, sort="-x")
)

glow = base.mark_bar(opacity=0.2, size=30, color="#2D9CFF")
bars = base.mark_bar(size=18, color="#2D9CFF")

st.altair_chart(glow + bars, use_container_width=True)


# =========================================================
# RANKING GARGALOS
# =========================================================
st.subheader("Ranking de Gargalos")

top = agg.sort_values("HORAS", ascending=False).head(5)

for i, row in top.iterrows():
    st.markdown(f"""
    <div class="tesla-card card-red">
        <div class="card-title">#{i+1} GARGALO</div>
        <div class="card-value">{row[col_F]}</div>
        <div class="card-title">{_fmt(row['HORAS'])} h</div>
    </div>
    """, unsafe_allow_html=True)
