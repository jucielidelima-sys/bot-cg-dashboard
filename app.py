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
# CSS - ESTILO INDUSTRIA 4.0 / TESLA
# =========================================================
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(180deg, #0f1116 0%, #151922 100%);
        color: #E8EDF7;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 1rem;
        max-width: 96%;
    }

    h1, h2, h3, h4, h5, h6 {
        color: #F5F7FB !important;
        font-weight: 700 !important;
        letter-spacing: 0.2px;
    }

    p, label, span, div {
        color: #D7DCE5;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #11151d 0%, #191f2b 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 18px;
        padding: 14px 16px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.18);
    }

    div[data-testid="stMetricLabel"] {
        color: #AEB7C6 !important;
        font-size: 0.9rem !important;
        font-weight: 600 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    .tesla-card {
        border-radius: 20px;
        padding: 18px 18px 14px 18px;
        margin-bottom: 14px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.22);
        border: 1px solid rgba(255,255,255,0.08);
        background: linear-gradient(135deg, rgba(255,255,255,0.06), rgba(255,255,255,0.03));
        backdrop-filter: blur(8px);
    }

    .card-green {
        border-left: 6px solid #14C38E;
    }

    .card-blue {
        border-left: 6px solid #2D9CFF;
    }

    .card-orange {
        border-left: 6px solid #FFB020;
    }

    .card-red {
        border-left: 6px solid #FF5A5F;
    }

    .card-purple {
        border-left: 6px solid #8B5CF6;
    }

    .card-title {
        color: #B6C0CF;
        font-size: 0.88rem;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        font-weight: 700;
    }

    .card-value {
        color: #FFFFFF;
        font-size: 2rem;
        font-weight: 800;
        line-height: 1.1;
    }

    .card-sub {
        color: #93A0B5;
        font-size: 0.84rem;
        margin-top: 6px;
    }

    .section-panel {
        border-radius: 20px;
        padding: 18px;
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.07);
        box-shadow: 0 8px 24px rgba(0,0,0,0.16);
        margin-bottom: 16px;
    }

    .small-note {
        color: #94A3B8;
        font-size: 0.82rem;
    }

    .stDataFrame, .stTable {
        background: rgba(255,255,255,0.03);
        border-radius: 16px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        background: rgba(255,255,255,0.04);
        color: #DCE3ED;
        padding: 10px 18px;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #2D9CFF, #8B5CF6) !important;
        color: white !important;
    }

    hr {
        border-color: rgba(255,255,255,0.08) !important;
    }
</style>
""", unsafe_allow_html=True)


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


def _find_col_exact_or_contains(df: pd.DataFrame, target: str) -> Optional[str]:
    if df is None or df.empty:
        return None

    target_norm = str(target).strip().lower()

    for c in df.columns:
        if str(c).strip().lower() == target_norm:
            return c

    for c in df.columns:
        if target_norm in str(c).strip().lower():
            return c

    return None


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
        return "#FF5A5F"
    if util_pct >= 85:
        return "#FFB020"
    return "#14C38E"


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


def _read_sheet_safe(xlsx_path: str, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        return pd.DataFrame()


def _card_html(title: str, value: str, subtitle: str = "", color_class: str = "card-blue") -> str:
    return f"""
    <div class="tesla-card {color_class}">
        <div class="card-title">{title}</div>
        <div class="card-value">{value}</div>
        <div class="card-sub">{subtitle}</div>
    </div>
    """


# =========================================================
# HEADER
# =========================================================
logo_img = _load_logo_image()
if logo_img is not None:
    st.image(logo_img, width=180)

st.title("Dashboard de Carga Máquina e Simulação de Mão de Obra")
st.caption("Base fixa no Excel da pasta • Visual industrial 4.0 • Simulação de cenários")


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

df_ind = _read_sheet_safe(ARQUIVO_EXCEL, "INDIRETOS")


# =========================================================
# MAPEAR COLUNAS
# =========================================================
col_C = _col_by_index(df0, 2)
col_F = _col_by_index(df0, 5)
col_J = _col_by_index(df0, 9)
col_R = _col_by_index(df0, 17)

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
# ABA INDIRETOS
# =========================================================
col_ind_setor = None
col_ind_moi = None
moi_total_fixo = 0.0
tabela_indiretos = pd.DataFrame()

if not df_ind.empty:
    col_ind_setor = _find_col_exact_or_contains(df_ind, "SETOR")
    col_ind_moi = _find_col_exact_or_contains(df_ind, "MOI")

    if col_ind_setor is not None and col_ind_moi is not None:
        tabela_indiretos = df_ind[[col_ind_setor, col_ind_moi]].copy()
        tabela_indiretos.columns = ["SETOR", "MOI"]

        tabela_indiretos["SETOR"] = tabela_indiretos["SETOR"].astype(str).str.strip()
        tabela_indiretos["MOI"] = tabela_indiretos["MOI"].apply(_to_float).fillna(0)

        tabela_indiretos = tabela_indiretos[
            tabela_indiretos["SETOR"].str.upper() != "TOTAL"
        ].copy()

        moi_total_fixo = float(tabela_indiretos["MOI"].sum())


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
    agg["cor"] = agg["util_pct"].apply(lambda x: _util_color(float(x)) if not np.isnan(x) else "#64748B")

    c1, c2 = st.columns([1.2, 1.0], gap="large")

    with c1:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.subheader("Carga por agrupamento")
        chart = alt.Chart(agg).mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6).encode(
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

        cap_line = alt.Chart(pd.DataFrame({"x": [cap_horas_efetivas]})).mark_rule(
            strokeDash=[6, 4],
            color="#FFFFFF"
        ).encode(x="x:Q")
        st.altair_chart(chart + cap_line, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.subheader("TAKT (soma)")
        chart2 = alt.Chart(agg).mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6).encode(
            x=alt.X("takt_h:Q", title="Horas TAKT"),
            y=alt.Y("_GRUPO_:N", sort="-x", title=""),
            color=alt.value("#2D9CFF"),
            tooltip=[
                alt.Tooltip("_GRUPO_:N", title="Grupo"),
                alt.Tooltip("takt_h:Q", title="TAKT somado", format=",.2f"),
            ],
        ).properties(height=min(600, 25 * max(8, len(agg))))
        st.altair_chart(chart2, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

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

    csv_maquina = detail.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Baixar dados carga máquina (CSV)",
        data=csv_maquina,
        file_name="carga_maquina_filtrada.csv",
        mime="text/csv",
        key="download_maquina"
    )


# =========================================================
# ABA 2 - MÃO DE OBRA
# =========================================================
with aba2:
    st.subheader("Simulação de Cenário - Mão de Obra")
    st.caption(
        "MOD = tempo individual da coluna G × quantidade planejada • "
        "MOI = mão de obra indireta fixa da aba INDIRETOS"
    )

    dias_mo = st.number_input(
        "Dias para cálculo da mão de obra direta",
        min_value=1.0,
        max_value=31.0,
        value=float(dias_uteis),
        step=1.0,
        key="dias_mo"
    )

    pessoas_disponiveis = st.number_input(
        "Pessoas disponíveis para comparação",
        min_value=0.0,
        max_value=10000.0,
        value=float(moi_total_fixo),
        step=1.0,
        key="pessoas_disponiveis"
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

    agg_mo["mod_pessoas"] = np.where(
        minutos_disponiveis_por_pessoa > 0,
        agg_mo["minutos_totais"] / minutos_disponiveis_por_pessoa,
        np.nan
    )
    agg_mo["mod_pessoas_arred"] = np.ceil(agg_mo["mod_pessoas"].fillna(0)).astype(int)

    total_min_mod = float(agg_mo["minutos_totais"].sum())
    total_mod = float(agg_mo["mod_pessoas"].sum())
    total_mod_arred = int(np.ceil(total_mod)) if np.isfinite(total_mod) else 0
    total_moi = float(moi_total_fixo)
    total_geral = total_mod + total_moi
    total_geral_arred = int(total_mod_arred + total_moi)
    saldo_pessoas = float(pessoas_disponiveis - total_geral)
    ocupacao_pessoas = (total_geral / pessoas_disponiveis * 100.0) if pessoas_disponiveis > 0 else np.nan

    cor_saldo = "card-green" if saldo_pessoas >= 0 else "card-red"
    cor_ocup = "card-green" if (not np.isnan(ocupacao_pessoas) and ocupacao_pessoas <= 85) else ("card-orange" if (not np.isnan(ocupacao_pessoas) and ocupacao_pessoas <= 100) else "card-red")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(_card_html(
            "MOD Necessária",
            _fmt_br(total_mod, 2),
            f"Base: {_fmt_br(MINUTOS_POR_PESSOA_DIA,0)} min/pessoa/dia",
            "card-blue"
        ), unsafe_allow_html=True)

    with c2:
        st.markdown(_card_html(
            "MOI Fixa",
            _fmt_br(total_moi, 0),
            "Lida da aba INDIRETOS",
            "card-purple"
        ), unsafe_allow_html=True)

    with c3:
        st.markdown(_card_html(
            "Total MOD + MOI",
            _fmt_br(total_geral, 2),
            f"Total arredondado: {total_geral_arred}",
            "card-orange"
        ), unsafe_allow_html=True)

    c4, c5, c6 = st.columns(3)
    with c4:
        st.markdown(_card_html(
            "Pessoas Disponíveis",
            _fmt_br(pessoas_disponiveis, 0),
            "Valor informado para comparação",
            "card-green"
        ), unsafe_allow_html=True)

    with c5:
        st.markdown(_card_html(
            "Saldo de Pessoas",
            _fmt_br(saldo_pessoas, 2),
            "Positivo = sobra • Negativo = falta",
            cor_saldo
        ), unsafe_allow_html=True)

    with c6:
        st.markdown(_card_html(
            "Ocupação da Equipe",
            f"{_fmt_br(ocupacao_pessoas,1)}%" if not np.isnan(ocupacao_pessoas) else "-",
            "Necessárias ÷ Disponíveis",
            cor_ocup
        ), unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="section-panel">
            <div class="small-note">
                Minutos totais MOD: <b>{_fmt_br(total_min_mod)}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
                Minutos disponíveis por pessoa no período: <b>{_fmt_br(minutos_disponiveis_por_pessoa)}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
                Dias do cenário: <b>{_fmt_br(dias_mo,0)}</b>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    g1, g2 = st.columns([1.15, 0.85], gap="large")

    with g1:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.subheader("MOD por Descrição CR")

        graf_mo = alt.Chart(agg_mo).mark_bar(
            cornerRadiusTopRight=6,
            cornerRadiusBottomRight=6
        ).encode(
            x=alt.X("mod_pessoas:Q", title="Pessoas necessárias (MOD)"),
            y=alt.Y("_DESC_CR_:N", sort="-x", title="Descrição CR"),
            color=alt.value("#2D9CFF"),
            tooltip=[
                alt.Tooltip("_DESC_CR_:N", title="Descrição CR"),
                alt.Tooltip("minutos_totais:Q", title="Minutos totais", format=",.2f"),
                alt.Tooltip("mod_pessoas:Q", title="MOD necessária", format=",.2f"),
                alt.Tooltip("mod_pessoas_arred:Q", title="MOD arredondada"),
            ],
        ).properties(height=min(720, 30 * max(8, len(agg_mo))))

        st.altair_chart(graf_mo, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.subheader("Necessárias x Disponíveis")

        resumo_pessoas = pd.DataFrame({
            "Tipo": ["MOD", "MOI", "Necessárias Total", "Disponíveis"],
            "Pessoas": [total_mod, total_moi, total_geral, pessoas_disponiveis],
            "Cor": ["#2D9CFF", "#8B5CF6", "#FFB020", "#14C38E"]
        })

        graf_resumo = alt.Chart(resumo_pessoas).mark_bar(cornerRadiusEnd=8).encode(
            x=alt.X("Pessoas:Q", title="Pessoas"),
            y=alt.Y("Tipo:N", title=""),
            color=alt.Color("Cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Tipo:N", title="Tipo"),
                alt.Tooltip("Pessoas:Q", title="Pessoas", format=",.2f"),
            ],
        ).properties(height=260)

        st.altair_chart(graf_resumo, use_container_width=True)

        comparativo = pd.DataFrame({
            "Categoria": ["Necessárias Total", "Disponíveis"],
            "Valor": [total_geral, pessoas_disponiveis]
        })

        graf_compare = alt.Chart(comparativo).mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8).encode(
            x=alt.X("Categoria:N", title=""),
            y=alt.Y("Valor:Q", title="Pessoas"),
            color=alt.condition(
                alt.datum.Categoria == "Disponíveis",
                alt.value("#14C38E"),
                alt.value("#FF5A5F")
            ),
            tooltip=[
                alt.Tooltip("Categoria:N", title="Categoria"),
                alt.Tooltip("Valor:Q", title="Pessoas", format=",.2f"),
            ],
        ).properties(height=260)

        st.subheader("Comparativo direto")
        st.altair_chart(graf_compare, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    t1, t2 = st.columns([1.2, 0.8], gap="large")

    with t1:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.markdown("### Tabela MOD por Descrição CR")
        tabela_mod = agg_mo.rename(columns={
            "_DESC_CR_": "DESCRIÇÃO CR",
            "minutos_totais": "MINUTOS_TOTAIS",
            "modelos": "MODELOS",
            "linhas": "LINHAS",
            "mod_pessoas": "MOD_PESSOAS",
            "mod_pessoas_arred": "MOD_PESSOAS_ARRED",
        }).copy()

        st.dataframe(tabela_mod, use_container_width=True, height=430)
        st.markdown("</div>", unsafe_allow_html=True)

    with t2:
        st.markdown('<div class="section-panel">', unsafe_allow_html=True)
        st.markdown("### MOI fixa (aba INDIRETOS)")
        if tabela_indiretos.empty:
            st.warning("Não encontrei dados válidos na aba INDIRETOS com colunas SETOR e MOI.")
        else:
            st.dataframe(tabela_indiretos, use_container_width=True, height=430)
        st.markdown("</div>", unsafe_allow_html=True)

    tabela_resumo_final = pd.DataFrame({
        "INDICADOR": [
            "MINUTOS_TOTAIS_MOD",
            "MOD_NECESSARIA",
            "MOD_ARREDONDADA",
            "MOI_FIXA",
            "TOTAL_MOD_MOI",
            "TOTAL_MOD_MOI_ARRED",
            "PESSOAS_DISPONIVEIS",
            "SALDO_PESSOAS",
            "OCUPACAO_EQUIPE_PCT"
        ],
        "VALOR": [
            total_min_mod,
            total_mod,
            total_mod_arred,
            total_moi,
            total_geral,
            total_geral_arred,
            pessoas_disponiveis,
            saldo_pessoas,
            ocupacao_pessoas
        ]
    })

    st.divider()
    d1, d2, d3 = st.columns(3)

    with d1:
        csv_mod = tabela_mod.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar MOD por Descrição CR (CSV)",
            data=csv_mod,
            file_name="mao_de_obra_direta_por_cr.csv",
            mime="text/csv",
            key="download_mod"
        )

    with d2:
        if not tabela_indiretos.empty:
            csv_moi = tabela_indiretos.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Baixar MOI fixa (CSV)",
                data=csv_moi,
                file_name="mao_de_obra_indireta_fixa.csv",
                mime="text/csv",
                key="download_moi"
            )

    with d3:
        csv_resumo = tabela_resumo_final.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar resumo total MOD + MOI (CSV)",
            data=csv_resumo,
            file_name="resumo_mao_de_obra_total.csv",
            mime="text/csv",
            key="download_resumo_mo"
        )
