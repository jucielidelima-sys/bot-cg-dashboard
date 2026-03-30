st.markdown("""
<style>
    html, body, [class*="css"]  {
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
            radial-gradient(circle at top left, rgba(45,156,255,0.10), transparent 26%),
            radial-gradient(circle at top right, rgba(139,92,246,0.10), transparent 24%),
            linear-gradient(180deg, #0a0d12 0%, #0f141d 50%, #131925 100%);
        color: #E8EDF7;
    }

    .metal-header {
        position: relative;
        overflow: hidden;
        border-radius: 22px;
        padding: 18px 24px;
        margin-top: 0rem !important;
        margin-bottom: 18px;
        background:
            linear-gradient(135deg, #616975 0%, #2e3540 18%, #8b949f 34%, #232a34 48%, #9aa4b0 64%, #2b323d 82%, #5f6874 100%);
        border: 1px solid rgba(255,255,255,0.20);
        box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.28),
            inset 0 -1px 0 rgba(0,0,0,0.25),
            0 10px 30px rgba(0,0,0,0.28);
    }

    .metal-header:before {
        content: "";
        position: absolute;
        inset: 0;
        background:
            linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.14) 30%, transparent 60%),
            repeating-linear-gradient(
                115deg,
                rgba(255,255,255,0.03) 0px,
                rgba(255,255,255,0.03) 2px,
                transparent 2px,
                transparent 10px
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
            0 0 10px rgba(255,255,255,0.14);
        margin: 0;
    }

    .metal-subtitle {
        margin-top: 6px;
        color: #E6ECF5;
        font-size: 0.93rem;
        font-weight: 500;
        letter-spacing: 0.3px;
    }

    section[data-testid="stSidebar"] {
        background:
            linear-gradient(180deg, rgba(255,255,255,0.04), rgba(255,255,255,0.01)),
            linear-gradient(180deg, #0f131b 0%, #171d29 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 18px;
        padding: 14px 16px;
        box-shadow:
            0 4px 16px rgba(0,0,0,0.18),
            inset 0 1px 0 rgba(255,255,255,0.03);
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
        text-shadow: 0 0 12px rgba(45,156,255,0.10);
    }

    .tesla-card {
        border-radius: 20px;
        padding: 18px 18px 14px 18px;
        margin-bottom: 14px;
        box-shadow:
            0 8px 24px rgba(0,0,0,0.24),
            0 0 18px rgba(45,156,255,0.06);
        border: 1px solid rgba(255,255,255,0.08);
        background:
            linear-gradient(135deg, rgba(255,255,255,0.07), rgba(255,255,255,0.03));
        backdrop-filter: blur(8px);
    }

    .card-green { border-left: 6px solid #14C38E; box-shadow: 0 0 18px rgba(20,195,142,0.12), 0 8px 24px rgba(0,0,0,0.24); }
    .card-blue { border-left: 6px solid #2D9CFF; box-shadow: 0 0 18px rgba(45,156,255,0.14), 0 8px 24px rgba(0,0,0,0.24); }
    .card-orange { border-left: 6px solid #FFB020; box-shadow: 0 0 18px rgba(255,176,32,0.12), 0 8px 24px rgba(0,0,0,0.24); }
    .card-red { border-left: 6px solid #FF5A5F; box-shadow: 0 0 18px rgba(255,90,95,0.14), 0 8px 24px rgba(0,0,0,0.24); }
    .card-purple { border-left: 6px solid #8B5CF6; box-shadow: 0 0 18px rgba(139,92,246,0.14), 0 8px 24px rgba(0,0,0,0.24); }

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
        text-shadow: 0 0 12px rgba(255,255,255,0.06);
    }

    .card-sub {
        color: #93A0B5;
        font-size: 0.82rem;
        margin-top: 8px;
    }

    .section-panel {
        border-radius: 20px;
        padding: 18px;
        background:
            linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.025));
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow:
            0 8px 24px rgba(0,0,0,0.20),
            inset 0 1px 0 rgba(255,255,255,0.04);
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
        border: 1px solid rgba(255,255,255,0.06);
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #2D9CFF, #8B5CF6) !important;
        color: white !important;
        box-shadow: 0 0 18px rgba(45,156,255,0.16);
    }

    .gargalo-panel {
        border-radius: 18px;
        padding: 14px 16px;
        margin-bottom: 10px;
        background: linear-gradient(90deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02));
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow: 0 6px 18px rgba(0,0,0,0.18);
    }

    .gargalo-top {
        border-left: 6px solid #FF5A5F;
        box-shadow: 0 0 16px rgba(255,90,95,0.12), 0 6px 18px rgba(0,0,0,0.18);
    }

    .gargalo-mid {
        border-left: 6px solid #FFB020;
        box-shadow: 0 0 16px rgba(255,176,32,0.10), 0 6px 18px rgba(0,0,0,0.18);
    }

    .gargalo-ok {
        border-left: 6px solid #14C38E;
        box-shadow: 0 0 16px rgba(20,195,142,0.10), 0 6px 18px rgba(0,0,0,0.18);
    }

    .gargalo-rank {
        font-size: 0.78rem;
        color: #9FB0C7;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        font-weight: 800;
    }

    .gargalo-name {
        font-size: 1rem;
        font-weight: 800;
        color: #F8FAFC;
        margin-top: 2px;
    }

    .gargalo-kpi {
        font-size: 1.35rem;
        font-weight: 900;
        color: #FFFFFF;
        margin-top: 4px;
    }

    .gargalo-sub {
        font-size: 0.8rem;
        color: #9BA8BC;
        margin-top: 4px;
    }

    hr {
        border-color: rgba(255,255,255,0.08) !important;
    }
</style>
""", unsafe_allow_html=True)
