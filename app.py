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
    page_title="Dashboard MES Industrial",
    layout="wide",
    initial_sidebar_state="collapsed",  # menu lateral recolhível / inicia recolhido
)


# =========================================================
# TEMA GLOBAL DOS GRÁFICOS
# =========================================================
def tema_dark_industrial():
    return {
        "config": {
            "background": "#0f1720",
            "view": {"stroke": "transparent"},
            "title": {
                "color": "#E8EDF7",
                "fontSize": 16,
                "fontWeight": 800
            },
            "axis": {
                "labelColor": "#AEB7C6",
                "titleColor": "#D7DCE5",
                "gridColor": "#263241",
                "domainColor": "#4A5668",
                "tickColor": "#4A5668"
            },
            "legend": {
                "labelColor": "#AEB7C6",
                "titleColor": "#D7DCE5"
            }
        }
    }


try:
    alt.themes.register("dark_industrial_mes", tema_dark_industrial)
except Exception:
    pass
alt.themes.enable("dark_industrial_mes")


# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>
html, body, [class*="css"] {
    margin: 0 !important;
    padding: 0 !important;
}

header[data-testid="stHeader"] {
    background: transparent !important;
    height: 0px !important;
}
header[data-testid="stHeader"] > div {
    height: 0px !important;
}

.block-container {
    padding-top: 0rem !important;
    padding-bottom: 1rem;
    max-width: 96%;
}

div[data-testid="stAppViewContainer"] > .main {
    padding-top: 0rem !important;
}
section.main > div {
    padding-top: 0rem !important;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(45,156,255,0.10), transparent 22%),
        radial-gradient(circle at top right, rgba(139,92,246,0.10), transparent 20%),
        radial-gradient(circle at bottom left, rgba(20,195,142,0.06), transparent 16%),
        linear-gradient(180deg, #090c11 0%, #0e141c 45%, #131b25 100%);
    color: #E8EDF7;
}

section[data-testid="stSidebar"] {
    background:
        linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.01)),
        linear-gradient(180deg, #0e1219 0%, #171e2a 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}

.metal-header {
    position: relative;
    overflow: hidden;
    border-radius: 24px;
    padding: 20px 26px;
    margin-top: 0rem !important;
    margin-bottom: 18px;
    background:
        linear-gradient(135deg, rgba(151,160,171,0.35), rgba(49,57,68,0.70) 18%, rgba(182,191,201,0.22) 34%, rgba(35,42,52,0.78) 52%, rgba(120,129,140,0.30) 72%, rgba(25,30,38,0.82) 100%);
    border: 1px solid rgba(255,255,255,0.16);
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,0.20),
        inset 0 -1px 0 rgba(0,0,0,0.28),
        0 12px 28px rgba(0,0,0,0.28),
        0 0 24px rgba(45,156,255,0.08);
    backdrop-filter: blur(12px);
}

.metal-header:before {
    content: "";
    position: absolute;
    inset: 0;
    background:
        linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.12) 32%, transparent 58%),
        repeating-linear-gradient(
            115deg,
            rgba(255,255,255,0.03) 0px,
            rgba(255,255,255,0.03) 2px,
            transparent 2px,
            transparent 12px
        );
    mix-blend-mode: screen;
    pointer-events: none;
}

.metal-title {
    font-size: 2rem;
    font-weight: 900;
    color: #F8FAFC;
    text-shadow:
        0 1px 0 rgba(0,0,0,0.35),
        0 0 10px rgba(255,255,255,0.10),
        0 0 18px rgba(45,156,255,0.06);
    margin: 0;
}

.metal-subtitle {
    margin-top: 6px;
    color: #E6ECF5;
    font-size: 0.93rem;
    font-weight: 500;
    letter-spacing: 0.3px;
}

.glass-panel {
    border-radius: 20px;
    padding: 18px;
    background: linear-gradient(180deg, rgba(255,255,255,0.065), rgba(255,255,255,0.025));
    border: 1px solid rgba(255,255,255,0.10);
    box-shadow:
        0 8px 24px rgba(0,0,0,0.22),
        inset 0 1px 0 rgba(255,255,255,0.05),
        0 0 18px rgba(45,156,255,0.04);
    backdrop-filter: blur(10px);
    margin-bottom: 16px;
}

.tesla-card {
    border-radius: 18px;
    padding: 16px 16px 12px 16px;
    margin-bottom: 12px;
    background: linear-gradient(135deg, rgba(255,255,255,0.07), rgba(255,255,255,0.03));
    border: 1px solid rgba(255,255,255,0.08);
    box-shadow:
        0 8px 22px rgba(0,0,0,0.22),
        0 0 18px rgba(45,156,255,0.05);
    backdrop-filter: blur(10px);
}

.card-green { border-left: 6px solid #14C38E; box-shadow: 0 0 16px rgba(20,195,142,0.12), 0 8px 22px rgba(0,0,0,0.22); }
.card-blue { border-left: 6px solid #2D9CFF; box-shadow: 0 0 16px rgba(45,156,255,0.14), 0 8px 22px rgba(0,0,0,0.22); }
.card-orange { border-left: 6px solid #FFB020; box-shadow: 0 0 16px rgba(255,176,32,0.12), 0 8px 22px rgba(0,0,0,0.22); }
.card-red { border-left: 6px solid #FF5A5F; box-shadow: 0 0 16px rgba(255,90,95,0.14), 0 8px 22px rgba(0,0,0,0.22); }
.card-purple { border-left: 6px solid #8B5CF6; box-shadow: 0 0 16px rgba(139,92,246,0.14), 0 8px 22px rgba(0,0,0,0.22); }

.card-title {
    color: #B6C0CF;
    font-size: 0.82rem;
    margin-bottom: 8px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    font-weight: 800;
}
.card-value {
    color: #FFFFFF;
    font-size: 2rem;
    font-weight: 900;
    line-height: 1.05;
}
.card-sub {
    color: #93A0B5;
    font-size: 0.82rem;
    margin-top: 8px;
}

.small-note {
    color: #94A3B8;
    font-size: 0.82rem;
}

.kpi-grid-title {
    color: #DDE6F3;
    font-size: 1rem;
    font-weight: 800;
    margin-bottom: 10px;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 12px;
    background: rgba(255,255,255,0.04);
    color: #DCE3ED;
    padding: 10px 18px;
    border: 1px solid rgba(255,255,255,0.06);
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(90deg, #2D9CFF, #8B5CF6) !important;
    color: white !important;
    box-shadow: 0 0 18px rgba(45,156,255,0.16);
}

.stDataFrame, .stTable {
    background: #FFFFFF !important;
    border-radius: 12px;
}
[data-testid="stDataFrame"] {
    background: #FFFFFF !important;
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
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.4px;
}
div[data-testid="stMetricValue"] {
    color: #FFFFFF !important;
    font-weight: 900 !important;
}

.rank-row {
    margin-bottom: 12px;
}
.rank-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 6px;
}
.rank-name {
    font-weight: 800;
    color: #F8FAFC;
    font-size: 0.95rem;
}
.rank-meta {
    color: #A5B1C2;
    font-size: 0.80rem;
    font-weight: 700;
}
.rank-track {
    width: 100%;
    height: 13px;
    background: rgba(255,255,255,0.06);
    border-radius: 999px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.08);
}
.rank-fill {
    height: 100%;
    border-radius: 999px;
    animation: growBar 1.2s ease-out forwards;
    transform-origin: left center;
    box-shadow: 0 0 12px currentColor;
}
.rank-red { background: linear-gradient(90deg, #FF5A5F, #FF7A7E); color: #FF5A5F; }
.rank-orange { background: linear-gradient(90deg, #FFB020, #FFD166); color: #FFB020; }
.rank-green { background: linear-gradient(90deg, #14C38E, #42E2B8); color: #14C38E; }
.rank-blue { background: linear-gradient(90deg, #2D9CFF, #62B7FF); color: #2D9CFF; }

@keyframes growBar {
    from { width: 0; opacity: 0.75; }
    to { opacity: 1; }
}

.neon-caption {
    color: #AFC2DA;
    font-size: 0.82rem;
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


def _rank_class(util_pct: float) -> str:
    if util_pct >= 100:
        return "rank-red"
    if util_pct >= 85:
        return "rank-orange"
    return "rank-green"


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


def _render_rank_bars(df_rank: pd.DataFrame, label_col: str, value_col: str, subtitle_col: Optional[str] = None, max_items: int = 5):
    if df_rank.empty:
        st.info("Sem dados para ranking.")
        return

    top = df_rank.head(max_items).copy()
    vmax = float(top[value_col].max()) if float(top[value_col].max()) > 0 else 1.0

    html_parts = []
    for idx, (_, row) in enumerate(top.iterrows(), start=1):
        nome = str(row[label_col])
        valor = float(row[value_col]) if pd.notna(row[value_col]) else 0.0
        pct = max(4.0, min(100.0, (valor / vmax) * 100.0))
        cls = _rank_class(valor)
        meta = str(row[subtitle_col]) if subtitle_col and subtitle_col in top.columns else _fmt_br(valor, 1)

        html_parts.append(f"""
        <div class="rank-row">
            <div class="rank-header">
                <div class="rank-name">#{idx} • {nome}</div>
                <div class="rank-meta">{meta}</div>
            </div>
            <div class="rank-track">
                <div class="rank-fill {cls}" style="width:{pct}%"></div>
            </div>
        </div>
        """)

    st.markdown("".join(html_parts), unsafe_allow_html=True)


# =========================================================
# HEADER
# =========================================================
logo_img = _load_logo_image()

col_logo, col_head = st.columns([0.12, 0.88])
with col_logo:
    if logo_img is not None:
        st.image(logo_img, width=135)

with col_head:
    st.markdown("""
    <div class="metal-header">
        <div class="metal-title">Dashboard MES Industrial</div>
        <div class="metal-subtitle">Tela executiva da fábrica • Sala de controle operacional • Simulação de cenários</div>
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# LOAD EXCEL FIXO
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
moi_total_fixo = 0.0
tabela_indiretos = pd.DataFrame()
tabela_indiretos_full = pd.DataFrame()

if not df_ind.empty:
    col_ind_setor = _find_col_exact_or_contains(df_ind, "SETOR")
    col_ind_moi = _find_col_exact_or_contains(df_ind, "MOI")
    col_ind_desc = _find_col_exact_or_contains(df_ind, "DESCRI")
    if col_ind_desc is None:
        col_ind_desc = _find_col_exact_or_contains(df_ind, "DESCRIÇÃO")

    cols_full = []
    for c in [col_ind_setor, col_ind_desc, col_ind_moi]:
        if c is not None and c not in cols_full:
            cols_full.append(c)

    if len(cols_full) > 0:
        tabela_indiretos_full = df_ind[cols_full].copy()

        rename_map = {}
        if col_ind_setor is not None:
            rename_map[col_ind_setor] = "SETOR"
        if col_ind_desc is not None:
            rename_map[col_ind_desc] = "DESCRIÇÃO"
        if col_ind_moi is not None:
            rename_map[col_ind_moi] = "MOI"

        tabela_indiretos_full = tabela_indiretos_full.rename(columns=rename_map)

        if "SETOR" in tabela_indiretos_full.columns:
            tabela_indiretos_full["SETOR"] = tabela_indiretos_full["SETOR"].astype(str).str.strip()

        if "DESCRIÇÃO" in tabela_indiretos_full.columns:
            tabela_indiretos_full["DESCRIÇÃO"] = tabela_indiretos_full["DESCRIÇÃO"].astype(str).str.strip()

        if "MOI" in tabela_indiretos_full.columns:
            tabela_indiretos_full["MOI"] = tabela_indiretos_full["MOI"].apply(_to_float).fillna(0.0)

        if "SETOR" in tabela_indiretos_full.columns:
            tabela_indiretos_full = tabela_indiretos_full[tabela_indiretos_full["SETOR"].str.upper() != "TOTAL"].copy()

    if col_ind_setor is not None and col_ind_moi is not None:
        tabela_indiretos = df_ind[[col_ind_setor, col_ind_moi]].copy()
        tabela_indiretos.columns = ["SETOR", "MOI"]
        tabela_indiretos["SETOR"] = tabela_indiretos["SETOR"].astype(str).str.strip()
        tabela_indiretos["MOI"] = tabela_indiretos["MOI"].apply(_to_float).fillna(0.0)
        tabela_indiretos = tabela_indiretos[tabela_indiretos["SETOR"].str.upper() != "TOTAL"].copy()
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

    st.divider()
    st.header("4) IA simples")

    crescimento_pct = st.slider("Crescimento previsto da demanda (%)", min_value=0, max_value=100, value=15, step=5)
    fator_previsao = 1.0 + (crescimento_pct / 100.0)


# =========================================================
# QUANTIDADE POR MODELO
# =========================================================
st.subheader("Quantidade por modelo")
st.caption("Digite a quantidade planejada por modelo. A carga usa TEMPO INDIVIDUAL da coluna G.")

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

# executivos
total_horas = float(df["HORAS_TRABALHADAS"].sum())
util_pct = (total_horas / cap_horas_efetivas * 100.0) if cap_horas_efetivas > 0 else np.nan
total_mod_min = float(df["CARGA_MIN"].sum())
total_mod_pessoas = (total_mod_min / (MINUTOS_POR_PESSOA_DIA * float(dias_uteis))) if dias_uteis > 0 else np.nan
total_pessoas_fabrica = float(total_mod_pessoas + moi_total_fixo) if not np.isnan(total_mod_pessoas) else float(moi_total_fixo)

# base agrupada executiva
tmp_exec = df.copy()
tmp_exec["_GRUPO_EXEC_"] = _col_series(df, col_F).astype(str).fillna("(vazio)")
agg_exec = tmp_exec.groupby("_GRUPO_EXEC_", dropna=False).agg(
    horas=("HORAS_TRABALHADAS", "sum"),
    n_cr=("_CR_CLEAN", pd.Series.nunique),
    qtd_modelos=(col_C, pd.Series.nunique),
).reset_index()
agg_exec["cap_prog_h"] = agg_exec["n_cr"].astype(float) * float(horas_periodo)
agg_exec["cap_efet_h"] = agg_exec["cap_prog_h"] * float(oee) * float(eff_mo)
agg_exec["util_pct"] = np.where(agg_exec["cap_efet_h"] > 0, agg_exec["horas"] / agg_exec["cap_efet_h"] * 100.0, np.nan)
agg_exec["horas_proj"] = agg_exec["horas"] * fator_previsao
agg_exec["util_proj_pct"] = np.where(agg_exec["cap_efet_h"] > 0, agg_exec["horas_proj"] / agg_exec["cap_efet_h"] * 100.0, np.nan)
agg_exec["score_ia"] = agg_exec["util_proj_pct"].fillna(0) + (agg_exec["qtd_modelos"].fillna(0) * 2.0)
agg_exec = agg_exec.sort_values("horas", ascending=False)
agg_exec_pred = agg_exec.sort_values(["score_ia", "util_proj_pct"], ascending=False)

gargalo_atual = str(agg_exec.iloc[0]["_GRUPO_EXEC_"]) if not agg_exec.empty else "-"
gargalo_atual_util = float(agg_exec.iloc[0]["util_pct"]) if (not agg_exec.empty and pd.notna(agg_exec.iloc[0]["util_pct"])) else np.nan
gargalo_previsto = str(agg_exec_pred.iloc[0]["_GRUPO_EXEC_"]) if not agg_exec_pred.empty else "-"
gargalo_previsto_util = float(agg_exec_pred.iloc[0]["util_proj_pct"]) if (not agg_exec_pred.empty and pd.notna(agg_exec_pred.iloc[0]["util_proj_pct"])) else np.nan


# =========================================================
# ABAS
# =========================================================
aba0, aba1, aba2, aba3 = st.tabs(["Resumo Executivo", "Carga Máquina", "Mão de Obra", "Indiretos"])


# =========================================================
# ABA 0 - RESUMO EXECUTIVO
# =========================================================
with aba0:
    st.subheader("Tela Inicial Executiva")
    st.caption("Visão geral da fábrica para acompanhamento rápido de capacidade, gargalos e mão de obra.")

    r1, r2, r3, r4 = st.columns(4)
    with r1:
        st.markdown(_card_html(
            "Carga Total",
            f"{_fmt_br(total_horas)} h",
            "Horas trabalhadas do cenário filtrado",
            "card-blue"
        ), unsafe_allow_html=True)
    with r2:
        st.markdown(_card_html(
            "Capacidade Efetiva",
            f"{_fmt_br(cap_horas_efetivas)} h",
            "Capacidade com OEE e eficiência MO",
            "card-green"
        ), unsafe_allow_html=True)
    with r3:
        st.markdown(_card_html(
            "Utilização Geral",
            f"{_fmt_br(util_pct,1)}%" if not np.isnan(util_pct) else "-",
            "Carga ÷ capacidade efetiva",
            "card-orange" if (not np.isnan(util_pct) and util_pct >= 85) else "card-green"
        ), unsafe_allow_html=True)
    with r4:
        st.markdown(_card_html(
            "Pessoas Totais",
            _fmt_br(total_pessoas_fabrica, 2),
            "MOD estimada + MOI fixa",
            "card-purple"
        ), unsafe_allow_html=True)

    r5, r6, r7, r8 = st.columns(4)
    with r5:
        st.markdown(_card_html(
            "MOD Estimada",
            _fmt_br(total_mod_pessoas, 2) if not np.isnan(total_mod_pessoas) else "-",
            f"Base: {int(MINUTOS_POR_PESSOA_DIA)} min/pessoa/dia",
            "card-blue"
        ), unsafe_allow_html=True)
    with r6:
        st.markdown(_card_html(
            "MOI Fixa",
            _fmt_br(moi_total_fixo, 0),
            "Lida da aba INDIRETOS",
            "card-purple"
        ), unsafe_allow_html=True)
    with r7:
        st.markdown(_card_html(
            "Gargalo Atual",
            gargalo_atual,
            f"Utilização: {_fmt_br(gargalo_atual_util,1)}%" if not np.isnan(gargalo_atual_util) else "Sem dados",
            "card-red" if (not np.isnan(gargalo_atual_util) and gargalo_atual_util >= 100) else "card-orange"
        ), unsafe_allow_html=True)
    with r8:
        st.markdown(_card_html(
            "Gargalo Previsto",
            gargalo_previsto,
            f"Projeção: {_fmt_br(gargalo_previsto_util,1)}%" if not np.isnan(gargalo_previsto_util) else "Sem dados",
            "card-red" if (not np.isnan(gargalo_previsto_util) and gargalo_previsto_util >= 100) else "card-orange"
        ), unsafe_allow_html=True)

    c1, c2 = st.columns([1.1, 0.9], gap="large")

    with c1:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Top 10 cargas por Descrição CR")

        top_exec = agg_exec.head(10).copy()
        top_exec["cor"] = top_exec["util_pct"].apply(lambda x: _util_color(float(x)) if pd.notna(x) else "#64748B")

        base_exec = alt.Chart(top_exec).encode(
            x=alt.X("horas:Q", title="Horas"),
            y=alt.Y("_GRUPO_EXEC_:N", sort="-x", title="")
        )

        glow_exec = base_exec.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            opacity=0.22,
            size=28
        ).encode(
            color=alt.Color("cor:N", scale=None, legend=None)
        )

        bars_exec = base_exec.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            size=16
        ).encode(
            color=alt.Color("cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("_GRUPO_EXEC_:N", title="Descrição CR"),
                alt.Tooltip("horas:Q", title="Horas", format=",.2f"),
                alt.Tooltip("util_pct:Q", title="Utilização", format=",.1f"),
            ],
        ).properties(height=360)

        st.altair_chart(glow_exec + bars_exec, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Ranking Executivo de Gargalos")

        rank_exec = agg_exec.sort_values("util_pct", ascending=False).copy()
        rank_exec["meta"] = rank_exec.apply(
            lambda r: f"{_fmt_br(float(r['util_pct']),1)}% • {_fmt_br(float(r['horas']),2)} h",
            axis=1
        )
        _render_rank_bars(rank_exec, "_GRUPO_EXEC_", "util_pct", "meta", max_items=8)
        st.markdown("</div>", unsafe_allow_html=True)

    c3, c4 = st.columns([1.0, 1.0], gap="large")

    with c3:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Painel Executivo de Capacidade")

        painel_exec = pd.DataFrame({
            "Indicador": ["Carga Total", "Cap. Programada", "Cap. Efetiva"],
            "Valor": [total_horas, cap_horas_programadas, cap_horas_efetivas],
            "Cor": ["#2D9CFF", "#8B5CF6", "#14C38E"]
        })

        chart_painel = alt.Chart(painel_exec).mark_bar(cornerRadiusEnd=8).encode(
            x=alt.X("Valor:Q", title="Horas"),
            y=alt.Y("Indicador:N", title=""),
            color=alt.Color("Cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Indicador:N", title="Indicador"),
                alt.Tooltip("Valor:Q", title="Horas", format=",.2f")
            ],
        ).properties(height=240)

        st.altair_chart(chart_painel, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Previsão Executiva")

        pred_exec = agg_exec_pred.head(5)[["_GRUPO_EXEC_", "util_proj_pct"]].copy()
        pred_exec["Cor"] = pred_exec["util_proj_pct"].apply(
            lambda x: "#FF5A5F" if x >= 100 else ("#FFB020" if x >= 85 else "#14C38E")
        )

        chart_pred_exec = alt.Chart(pred_exec).mark_bar(cornerRadiusEnd=8).encode(
            x=alt.X("util_proj_pct:Q", title="Utilização projetada (%)"),
            y=alt.Y("_GRUPO_EXEC_:N", sort="-x", title=""),
            color=alt.Color("Cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("_GRUPO_EXEC_:N", title="Descrição CR"),
                alt.Tooltip("util_proj_pct:Q", title="Utilização projetada", format=",.1f"),
            ],
        ).properties(height=240)

        st.altair_chart(chart_pred_exec, use_container_width=True)
        st.markdown(
            f"""
            <div class="small-note">
                Cenário de projeção ativo: <b>+{crescimento_pct}%</b> na demanda.<br>
                Gargalo previsto principal: <b>{gargalo_previsto}</b>.
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# ABA 1 - CARGA MÁQUINA
# =========================================================
with aba1:
    st.subheader("Carga Máquina")

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
        qtd_modelos=(col_C, pd.Series.nunique),
    ).reset_index()

    agg = agg.sort_values("horas", ascending=False)
    agg["cap_prog_h"] = agg["n_cr"].astype(float) * float(horas_periodo)
    agg["cap_efet_h"] = agg["cap_prog_h"] * float(oee) * float(eff_mo)
    agg["util_pct"] = np.where(agg["cap_efet_h"] > 0, agg["horas"] / agg["cap_efet_h"] * 100.0, np.nan)
    agg["cor"] = agg["util_pct"].apply(lambda x: _util_color(float(x)) if not np.isnan(x) else "#64748B")

    agg["horas_proj"] = agg["horas"] * fator_previsao
    agg["util_proj_pct"] = np.where(agg["cap_efet_h"] > 0, agg["horas_proj"] / agg["cap_efet_h"] * 100.0, np.nan)
    agg["score_ia"] = agg["util_proj_pct"].fillna(0) + (agg["qtd_modelos"].fillna(0) * 2.0)
    agg_pred = agg.sort_values(["score_ia", "util_proj_pct"], ascending=False).copy()

    c1, c2 = st.columns([1.15, 0.85], gap="large")

    with c1:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Carga por agrupamento")
        st.markdown('<div class="neon-caption">Glow neon com sobreposição para efeito visual industrial.</div>', unsafe_allow_html=True)

        base = alt.Chart(agg).encode(
            x=alt.X("horas:Q", title="Horas (carga)"),
            y=alt.Y("_GRUPO_:N", sort="-x", title="")
        )

        glow = base.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            opacity=0.22,
            size=28
        ).encode(
            color=alt.Color("cor:N", scale=None, legend=None)
        )

        bars = base.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            size=16
        ).encode(
            color=alt.Color("cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("_GRUPO_:N", title="Grupo"),
                alt.Tooltip("horas:Q", title="Horas", format=",.2f"),
                alt.Tooltip("cap_efet_h:Q", title="Cap. efetiva", format=",.2f"),
                alt.Tooltip("util_pct:Q", title="Utilização %", format=",.1f"),
            ],
        ).properties(height=min(620, 28 * max(8, len(agg))))

        cap_line = alt.Chart(pd.DataFrame({"x": [cap_horas_efetivas]})).mark_rule(
            strokeDash=[6, 4],
            color="#E2E8F0",
            size=2
        ).encode(x="x:Q")

        st.altair_chart(glow + bars + cap_line, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Ranking de Gargalos - MES")

        rank_df = agg.sort_values("util_pct", ascending=False).copy()
        rank_df["meta"] = rank_df.apply(
            lambda r: f"{_fmt_br(float(r['util_pct']),1)}% • {_fmt_br(float(r['horas']),2)} h",
            axis=1
        )
        _render_rank_bars(rank_df, "_GRUPO_", "util_pct", "meta", max_items=6)
        st.markdown("</div>", unsafe_allow_html=True)

    c3, c4 = st.columns([1.0, 1.0], gap="large")

    with c3:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("TAKT (soma)")

        base2 = alt.Chart(agg).encode(
            x=alt.X("takt_h:Q", title="Horas TAKT"),
            y=alt.Y("_GRUPO_:N", sort="-x", title="")
        )

        glow2 = base2.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            opacity=0.20,
            size=26,
            color="#2D9CFF"
        )

        bars2 = base2.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            size=16,
            color="#2D9CFF"
        ).encode(
            tooltip=[
                alt.Tooltip("_GRUPO_:N", title="Grupo"),
                alt.Tooltip("takt_h:Q", title="TAKT somado", format=",.2f"),
            ],
        ).properties(height=min(520, 26 * max(8, len(agg))))

        st.altair_chart(glow2 + bars2, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Previsão de Gargalo (IA simples)")

        if agg_pred.empty:
            st.info("Sem dados para projeção.")
        else:
            pred_top = agg_pred.iloc[0]
            pred_nome = str(pred_top["_GRUPO_"])
            pred_util = float(pred_top["util_proj_pct"]) if pd.notna(pred_top["util_proj_pct"]) else 0.0
            pred_carga = float(pred_top["horas_proj"]) if pd.notna(pred_top["horas_proj"]) else 0.0
            pred_cap = float(pred_top["cap_efet_h"]) if pd.notna(pred_top["cap_efet_h"]) else 0.0

            st.markdown(_card_html(
                "Gargalo Previsto",
                pred_nome,
                f"Cenário: +{crescimento_pct}% demanda",
                "card-red" if pred_util >= 100 else ("card-orange" if pred_util >= 85 else "card-green")
            ), unsafe_allow_html=True)

            st.markdown(_card_html(
                "Utilização Projetada",
                f"{_fmt_br(pred_util,1)}%",
                f"Carga projetada: {_fmt_br(pred_carga)} h • Cap.: {_fmt_br(pred_cap)} h",
                "card-blue"
            ), unsafe_allow_html=True)

            pred_small = agg_pred[["_GRUPO_", "util_proj_pct"]].copy().head(5)
            pred_small["Cor"] = pred_small["util_proj_pct"].apply(
                lambda x: "#FF5A5F" if x >= 100 else ("#FFB020" if x >= 85 else "#14C38E")
            )

            chart_pred = alt.Chart(pred_small).mark_bar(cornerRadiusEnd=8).encode(
                x=alt.X("util_proj_pct:Q", title="Utilização projetada (%)"),
                y=alt.Y("_GRUPO_:N", sort="-x", title=""),
                color=alt.Color("Cor:N", scale=None, legend=None),
                tooltip=[
                    alt.Tooltip("_GRUPO_:N", title="Grupo"),
                    alt.Tooltip("util_proj_pct:Q", title="Utilização projetada", format=",.1f"),
                ],
            ).properties(height=240)

            st.altair_chart(chart_pred, use_container_width=True)
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
    if np.isnan(ocupacao_pessoas):
        cor_ocup = "card-blue"
    elif ocupacao_pessoas <= 85:
        cor_ocup = "card-green"
    elif ocupacao_pessoas <= 100:
        cor_ocup = "card-orange"
    else:
        cor_ocup = "card-red"

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
        <div class="glass-panel">
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
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("MOD por Descrição CR")

        base3 = alt.Chart(agg_mo).encode(
            x=alt.X("mod_pessoas:Q", title="Pessoas necessárias (MOD)"),
            y=alt.Y("_DESC_CR_:N", sort="-x", title="Descrição CR")
        )

        glow3 = base3.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            opacity=0.20,
            size=26,
            color="#2D9CFF"
        )

        bars3 = base3.mark_bar(
            cornerRadiusTopRight=7,
            cornerRadiusBottomRight=7,
            size=16,
            color="#2D9CFF"
        ).encode(
            tooltip=[
                alt.Tooltip("_DESC_CR_:N", title="Descrição CR"),
                alt.Tooltip("minutos_totais:Q", title="Minutos totais", format=",.2f"),
                alt.Tooltip("mod_pessoas:Q", title="MOD necessária", format=",.2f"),
                alt.Tooltip("mod_pessoas_arred:Q", title="MOD arredondada"),
            ],
        ).properties(height=min(720, 30 * max(8, len(agg_mo))))

        st.altair_chart(glow3 + bars3, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.subheader("Necessárias x Disponíveis")

        resumo_pessoas = pd.DataFrame({
            "Tipo": ["MOD", "MOI", "Necessárias Total", "Disponíveis"],
            "Pessoas": [total_mod, total_moi, total_geral, pessoas_disponiveis],
            "Cor": ["#2D9CFF", "#8B5CF6", "#FFB020", "#14C38E"]
        })

        chart_neon_h = alt.Chart(resumo_pessoas).encode(
            x=alt.X("Pessoas:Q", title="Pessoas"),
            y=alt.Y("Tipo:N", title="")
        )

        glow_h = chart_neon_h.mark_bar(cornerRadiusEnd=8, opacity=0.20, size=26).encode(
            color=alt.Color("Cor:N", scale=None, legend=None)
        )

        bars_h = chart_neon_h.mark_bar(cornerRadiusEnd=8, size=18).encode(
            color=alt.Color("Cor:N", scale=None, legend=None),
            tooltip=[
                alt.Tooltip("Tipo:N", title="Tipo"),
                alt.Tooltip("Pessoas:Q", title="Pessoas", format=",.2f"),
            ],
        ).properties(height=260)

        st.altair_chart(glow_h + bars_h, use_container_width=True)

        comp_df = pd.DataFrame({
            "nome": ["Necessárias", "Disponíveis"],
            "valor": [total_geral, pessoas_disponiveis]
        })
        comp_df["meta"] = comp_df["valor"].apply(lambda x: f"{_fmt_br(float(x),2)} pessoas")

        st.subheader("Ranking comparativo - MES")
        _render_rank_bars(
            comp_df.sort_values("valor", ascending=False),
            "nome",
            "valor",
            "meta",
            max_items=2
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    t1, t2 = st.columns([1.2, 0.8], gap="large")

    with t1:
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
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
        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
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


# =========================================================
# ABA 3 - INDIRETOS
# =========================================================
with aba3:
    st.subheader("Descrição dos Indiretos")
    st.caption("Visualização da aba INDIRETOS com descrição das atividades e quantitativo de MOI.")

    if tabela_indiretos_full.empty:
        st.warning("Não encontrei dados válidos na aba INDIRETOS.")
    else:
        c1, c2, c3 = st.columns(3)

        total_moi_ind = float(tabela_indiretos_full["MOI"].sum()) if "MOI" in tabela_indiretos_full.columns else 0.0
        qtd_setores_ind = int(tabela_indiretos_full["SETOR"].nunique()) if "SETOR" in tabela_indiretos_full.columns else 0
        qtd_descr_ind = int(tabela_indiretos_full["DESCRIÇÃO"].nunique()) if "DESCRIÇÃO" in tabela_indiretos_full.columns else 0

        with c1:
            st.markdown(_card_html(
                "MOI Total",
                _fmt_br(total_moi_ind, 0),
                "Soma total dos indiretos",
                "card-purple"
            ), unsafe_allow_html=True)

        with c2:
            st.markdown(_card_html(
                "Setores",
                str(qtd_setores_ind),
                "Quantidade de setores cadastrados",
                "card-blue"
            ), unsafe_allow_html=True)

        with c3:
            st.markdown(_card_html(
                "Descrições",
                str(qtd_descr_ind),
                "Atividades indiretas cadastradas",
                "card-green"
            ), unsafe_allow_html=True)

        st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
        st.markdown("### Tabela completa de indiretos")
        st.dataframe(tabela_indiretos_full, use_container_width=True, height=500)
        st.markdown("</div>", unsafe_allow_html=True)

        if "DESCRIÇÃO" in tabela_indiretos_full.columns and "MOI" in tabela_indiretos_full.columns:
            agg_ind_desc = (
                tabela_indiretos_full.groupby("DESCRIÇÃO", dropna=False)["MOI"]
                .sum()
                .reset_index()
                .sort_values("MOI", ascending=False)
            )

            c4, c5 = st.columns([1.1, 0.9], gap="large")

            with c4:
                st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
                st.subheader("MOI por descrição")

                base_ind = alt.Chart(agg_ind_desc).encode(
                    x=alt.X("MOI:Q", title="MOI"),
                    y=alt.Y("DESCRIÇÃO:N", sort="-x", title="")
                )

                glow_ind = base_ind.mark_bar(
                    cornerRadiusTopRight=7,
                    cornerRadiusBottomRight=7,
                    opacity=0.20,
                    size=26,
                    color="#8B5CF6"
                )

                bars_ind = base_ind.mark_bar(
                    cornerRadiusTopRight=7,
                    cornerRadiusBottomRight=7,
                    size=16,
                    color="#8B5CF6"
                ).encode(
                    tooltip=[
                        alt.Tooltip("DESCRIÇÃO:N", title="Descrição"),
                        alt.Tooltip("MOI:Q", title="MOI", format=",.0f"),
                    ],
                ).properties(height=min(700, 28 * max(8, len(agg_ind_desc))))

                st.altair_chart(glow_ind + bars_ind, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)

            with c5:
                st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
                st.subheader("Ranking de Indiretos - MES")

                rank_ind = agg_ind_desc.copy()
                rank_ind["meta"] = rank_ind["MOI"].apply(lambda x: f"{_fmt_br(float(x),0)} MOI")
                _render_rank_bars(rank_ind, "DESCRIÇÃO", "MOI", "meta", max_items=8)
                st.markdown("</div>", unsafe_allow_html=True)

        csv_ind = tabela_indiretos_full.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Baixar indiretos completos (CSV)",
            data=csv_ind,
            file_name="indiretos_completo.csv",
            mime="text/csv",
            key="download_indiretos_full"
        )
