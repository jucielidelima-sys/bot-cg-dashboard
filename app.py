import os
import re
from typing import Optional, List, Dict

import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image
import altair as alt


# =========================================================
# CONFIG
# =========================================================
ARQUIVO_EXCEL = "CG BOT PY.xlsx"
MINUTOS_POR_PESSOA_DIA = 500.0


# =========================================================
# STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Dashboard Industrial",
    layout="wide"
)


# =========================================================
# CSS INDUSTRIAL TESLA
# =========================================================
st.markdown("""
<style>
html, body {
    margin:0;
    padding:0;
}

header {visibility:hidden;}

.block-container {
    padding-top: 0rem;
}

.stApp {
    background: linear-gradient(180deg,#0b0f16,#111826,#141d2b);
    color:white;
}

/* HEADER */
.metal-header {
    padding:18px;
    border-radius:20px;
    background: linear-gradient(135deg,#5c6672,#1f2630,#7c8794);
    margin-bottom:20px;
}

/* CARDS */
.card {
    border-radius:14px;
    padding:16px;
    background:rgba(255,255,255,0.05);
    box-shadow:0 0 10px rgba(45,156,255,0.2);
    margin-bottom:10px;
}

/* RANK */
.rank-bar {
    height:14px;
    border-radius:10px;
    margin-bottom:8px;
    background:linear-gradient(90deg,#2D9CFF,#8B5CF6);
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# HELPERS
# =========================================================
def _to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except:
        return 0


def _col(df, i):
    return df.columns[i]


# =========================================================
# LOAD
# =========================================================
df = pd.read_excel(ARQUIVO_EXCEL)
df_ind = pd.read_excel(ARQUIVO_EXCEL, sheet_name="INDIRETOS")


col_C = _col(df, 2)
col_F = _col(df, 5)
col_G = _col(df, 6)

df["TEMPO"] = df[col_G].apply(_to_float)


# =========================================================
# HEADER
# =========================================================
col1, col2 = st.columns([0.1,0.9])

with col1:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=100)

with col2:
    st.markdown("""
    <div class="metal-header">
    <h2>Dashboard Industrial 4.0</h2>
    <p>Sala de controle MES</p>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    modelos = st.multiselect("Modelo", df[col_C].unique())
    cr = st.multiselect("CR", df[col_F].unique())

    crescimento = st.slider("Crescimento %",0,100,10)


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
st.subheader("Quantidade")

qty = {}
for m in modelos:
    qty[m] = st.number_input(m,0,100000,100)

df["QTD"] = df[col_C].map(qty).fillna(0)
df["MIN"] = df["TEMPO"] * df["QTD"]
df["HORAS"] = df["MIN"] / 60


# =========================================================
# ABAS
# =========================================================
tab1, tab2, tab3 = st.tabs(["Carga Máquina","Mão de Obra","Indiretos"])


# =========================================================
# ABA 1
# =========================================================
with tab1:

    total = df["HORAS"].sum()

    st.metric("Horas Totais", round(total,2))

    agg = df.groupby(col_F)["HORAS"].sum().reset_index()

    chart = alt.Chart(agg).mark_bar().encode(
        x="HORAS",
        y=col_F
    )
    st.altair_chart(chart, use_container_width=True)

    # RANK
    st.subheader("Gargalos")

    top = agg.sort_values("HORAS",ascending=False).head(5)

    for i,row in top.iterrows():
        pct = (row["HORAS"]/top["HORAS"].max())*100
        st.write(row[col_F])
        st.markdown(f"<div class='rank-bar' style='width:{pct}%'></div>", unsafe_allow_html=True)


# =========================================================
# ABA 2
# =========================================================
with tab2:

    total_min = df["MIN"].sum()
    mod = total_min / 500

    moi = df_ind["MOI"].sum()

    total_pessoas = mod + moi

    st.metric("MOD", round(mod,2))
    st.metric("MOI", round(moi,2))
    st.metric("Total", round(total_pessoas,2))

    comp = pd.DataFrame({
        "Tipo":["MOD","MOI"],
        "Valor":[mod,moi]
    })

    chart2 = alt.Chart(comp).mark_bar().encode(
        x="Tipo",
        y="Valor"
    )
    st.altair_chart(chart2, use_container_width=True)


# =========================================================
# ABA 3 - INDIRETOS
# =========================================================
with tab3:

    st.subheader("Indiretos")

    st.dataframe(df_ind)

    if "DESCRIÇÃO" in df_ind.columns:
        agg_ind = df_ind.groupby("DESCRIÇÃO")["MOI"].sum().reset_index()

        chart3 = alt.Chart(agg_ind).mark_bar().encode(
            x="MOI",
            y="DESCRIÇÃO"
        )
        st.altair_chart(chart3, use_container_width=True)
