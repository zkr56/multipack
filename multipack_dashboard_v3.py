"""
╔══════════════════════════════════════════════════════════════╗
║         MULTIPACK SA – TABLEAU DE BORD DIRECTION             ║
║     Zone Industrielle de Yopougon, Abidjan – Côte d'Ivoire   ║
║   Graphiques identiques aux fichiers Excel (couleurs & types) ║
╚══════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# ══════════════════════════════════════════════
# 0. CONFIG
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="MULTIPACK SA – Direction",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Couleurs extraites des fichiers Excel ──────────────────────
# Graphiques Application.xlsm  → #6172F3 (bleu-violet principal)
#                                #C5CBFB (bleu clair secondaire)
#                                #FF6BA7 (rose accent)
#                                #3C41CD (bleu foncé)
# Graphiques Gestion_de_stock  → #51459E (violet foncé)
#                                standard palette Excel (pie)
# Fond des graphiques Excel    → BLANC (#FFFFFF)
# Fond général du dashboard    → gris clair (#F4F6FB)

XL_BLUE      = "#6172F3"   # couleur principale Excel
XL_BLUE_LIGHT= "#C5CBFB"   # bleu clair Excel
XL_PINK      = "#FF6BA7"   # rose accent Excel
XL_BLUE_DARK = "#3C41CD"   # bleu foncé Excel
XL_PURPLE    = "#51459E"   # violet stock Excel
XL_GREEN     = "#20BF6B"   # vert (payé)
XL_RED       = "#E74C3C"   # rouge (impayé/alerte)
XL_ORANGE    = "#F39C12"   # orange
XL_TEAL      = "#00B894"   # vert-bleu
XL_CYAN      = "#0ABDE3"   # cyan

# Palette pie/donut standard Excel (ordre des couleurs auto Excel)
EXCEL_PIE_COLORS = [
    "#4472C4",  # bleu Excel accent5
    "#ED7D31",  # orange Excel accent2
    "#A5A5A5",  # gris Excel accent3
    "#FFC000",  # or Excel accent4
    "#5B9BD5",  # bleu clair Excel accent1
    "#70AD47",  # vert Excel accent6
    "#264478",  # bleu foncé
    "#9E480E",  # marron
]

# Palette area/bar principale (style fichiers Excel)
EXCEL_PALETTE = [XL_BLUE, XL_PINK, XL_PURPLE, XL_BLUE_DARK,
                 XL_BLUE_LIGHT, "#F7B731", "#20BF6B", "#0ABDE3"]

# Fond BLANC pour tous les graphiques (comme Excel)
CHART_BG = "#FFFFFF"
CHART_PLOT_BG = "#FFFFFF"
GRID_COLOR = "#E8ECF0"
AXIS_COLOR = "#8A94A6"
FONT_COLOR = "#2D3748"

# ── Couleurs UI dashboard ──────────────────────────────────────
UI = {
    "bg":        "#F4F6FB",   # fond gris très clair
    "card":      "#FFFFFF",   # cartes blanches
    "border":    "#E2E8F0",
    "sidebar":   "#1E2A4A",   # sidebar bleu marine foncé
    "header_bg": "#1E2A4A",
    "text":      "#2D3748",
    "muted":     "#718096",
    "rouge":     "#E53E3E",
    "vert":      "#38A169",
    "or":        "#D69E2E",
    "bleu":      XL_BLUE,
}

# ══════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

*, *::before, *::after {{ box-sizing: border-box; }}

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background-color: {UI['bg']};
    color: {UI['text']};
}}

/* FOND GLOBAL */
.main .block-container {{
    background-color: {UI['bg']};
    padding: 1.2rem 2rem 2rem;
    max-width: 1600px;
}}

/* SIDEBAR */
[data-testid="stSidebar"] {{
    background: {UI['sidebar']} !important;
    border-right: none;
}}
[data-testid="stSidebar"] * {{ color: #E2E8F0 !important; }}
[data-testid="stSidebar"] .stRadio > label {{
    color: #94A3B8 !important;
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.08em !important;
}}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {{
    background: rgba(255,255,255,0.04);
    border-radius: 8px;
    padding: 9px 14px !important;
    margin: 3px 0 !important;
    border: 1px solid transparent;
    transition: all 0.15s;
    font-size: 0.84rem !important;
    font-weight: 500 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    color: #CBD5E1 !important;
}}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {{
    background: rgba(97,114,243,0.18) !important;
    border-color: rgba(97,114,243,0.4) !important;
}}
[data-testid="stSidebar"] [data-testid="stMultiSelect"] *,
[data-testid="stSidebar"] [data-testid="stSelectbox"] * {{
    color: {UI['text']} !important;
}}

/* HEADER BANNER */
.mp-header {{
    background: linear-gradient(135deg, #1E2A4A 0%, #2D3A6B 50%, {XL_BLUE} 100%);
    border-radius: 14px;
    padding: 24px 32px;
    margin-bottom: 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    box-shadow: 0 4px 20px rgba(97,114,243,0.25);
}}
.mp-header h1 {{
    font-size: 1.6rem; font-weight: 800; color: white;
    margin: 0 0 3px; letter-spacing: -0.02em;
}}
.mp-header .sub {{ color: rgba(255,255,255,0.70); font-size: 0.82rem; }}
.mp-header .date-badge {{
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 10px; padding: 10px 18px;
    text-align: center; color: white;
    min-width: 130px;
}}
.mp-header .date-badge .day {{
    font-size: 1.5rem; font-weight: 800;
    display: block; line-height: 1.1;
}}

/* KPI CARDS */
.kpi-card {{
    background: {UI['card']};
    border-radius: 12px;
    padding: 18px 20px;
    border: 1px solid {UI['border']};
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    position: relative;
    overflow: hidden;
    transition: box-shadow 0.2s;
    height: 100%;
}}
.kpi-card:hover {{ box-shadow: 0 4px 16px rgba(0,0,0,0.10); }}
.kpi-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 4px;
    border-radius: 12px 12px 0 0;
}}
.kpi-card.bleu::before   {{ background: {XL_BLUE}; }}
.kpi-card.vert::before   {{ background: {UI['vert']}; }}
.kpi-card.rouge::before  {{ background: {UI['rouge']}; }}
.kpi-card.or::before     {{ background: {UI['or']}; }}
.kpi-card.pink::before   {{ background: {XL_PINK}; }}
.kpi-card.purple::before {{ background: {XL_PURPLE}; }}

.kpi-card .icon {{ font-size: 1.4rem; margin-bottom: 8px; display: block; }}
.kpi-card .lbl  {{ font-size: 0.70rem; font-weight: 600; color: {UI['muted']}; text-transform: uppercase; letter-spacing: 0.07em; }}
.kpi-card .val  {{ font-size: 1.55rem; font-weight: 800; color: {UI['text']}; margin: 5px 0 3px; line-height: 1.15; }}
.kpi-card .sub  {{ font-size: 0.74rem; color: {UI['muted']}; }}
.kpi-card .dpos {{ color: {UI['vert']}; font-weight: 700; }}
.kpi-card .dneg {{ color: {UI['rouge']}; font-weight: 700; }}

/* SECTION TITLES */
.sec-title {{
    font-size: 0.82rem; font-weight: 700; color: {UI['muted']};
    text-transform: uppercase; letter-spacing: 0.09em;
    padding-bottom: 8px; margin: 20px 0 12px;
    border-bottom: 2px solid {UI['border']};
    display: flex; align-items: center; gap: 7px;
}}
.sec-title .accent {{ color: {XL_BLUE}; font-size: 1rem; }}

/* CHART WRAPPER */
.chart-wrap {{
    background: {CHART_BG};
    border-radius: 12px;
    border: 1px solid {UI['border']};
    padding: 16px 18px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    height: 100%;
}}
.chart-wrap .ctitle {{ font-size: 0.85rem; font-weight: 600; color: {UI['text']}; margin-bottom: 2px; }}
.chart-wrap .csub   {{ font-size: 0.71rem; color: {UI['muted']}; margin-bottom: 12px; }}

/* ALERT CARDS */
.alert-r {{
    background: #FFF5F5; border-left: 4px solid {UI['rouge']};
    border-radius: 8px; padding: 11px 14px; margin: 5px 0;
    font-size: 0.82rem; color: #742A2A;
    display: flex; align-items: flex-start; gap: 9px;
}}
.alert-y {{
    background: #FFFFF0; border-left: 4px solid {UI['or']};
    border-radius: 8px; padding: 11px 14px; margin: 5px 0;
    font-size: 0.82rem; color: #744210;
    display: flex; align-items: flex-start; gap: 9px;
}}
.alert-g {{
    background: #F0FFF4; border-left: 4px solid {UI['vert']};
    border-radius: 8px; padding: 11px 14px; margin: 5px 0;
    font-size: 0.82rem; color: #1A4731;
    display: flex; align-items: flex-start; gap: 9px;
}}

/* PROGRESS */
.prog-wrap {{ margin: 7px 0; }}
.prog-labels {{ display: flex; justify-content: space-between; font-size: 0.73rem; color: {UI['muted']}; margin-bottom: 4px; }}
.prog-bar {{ height: 7px; background: {UI['border']}; border-radius: 4px; overflow: hidden; }}
.prog-fill {{ height: 100%; border-radius: 4px; }}

/* STAT PILL */
.stat-row {{
    background: {UI['card']}; border: 1px solid {UI['border']};
    border-radius: 10px; padding: 12px 16px;
    display: flex; justify-content: space-between; align-items: center;
    margin: 5px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}}
.stat-row .sl {{ font-size: 0.8rem; color: {UI['muted']}; }}
.stat-row .sv {{ font-size: 1rem; font-weight: 700; color: {UI['text']}; }}

#MainMenu, footer {{ visibility: hidden; }}
.stDeployButton {{ display: none; }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# THEME PLOTLY — FOND BLANC COMME EXCEL
# ══════════════════════════════════════════════
def apply_excel_style(fig, height=320, show_legend=True):
    """Applique le style Excel aux graphiques Plotly : fond blanc, grille légère."""
    fig.update_layout(
        paper_bgcolor=CHART_BG,
        plot_bgcolor=CHART_PLOT_BG,
        font=dict(family="Inter, sans-serif", color=FONT_COLOR, size=11),
        height=height,
        margin=dict(t=10, b=35, l=10, r=10),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            font=dict(color=FONT_COLOR, size=10),
            orientation="h",
            yanchor="bottom", y=1.02,
            xanchor="right", x=1,
        ) if show_legend else dict(visible=False),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_family="Inter",
            bordercolor=UI["border"],
        ),
    )
    fig.update_xaxes(
        gridcolor=GRID_COLOR, showgrid=True, zeroline=False,
        linecolor=UI["border"], tickcolor=AXIS_COLOR,
        tickfont=dict(color=AXIS_COLOR, size=10),
    )
    fig.update_yaxes(
        gridcolor=GRID_COLOR, showgrid=True, zeroline=False,
        linecolor=UI["border"], tickcolor=AXIS_COLOR,
        tickfont=dict(color=AXIS_COLOR, size=10),
    )
    return fig


def section(emoji, title):
    st.markdown(f'<div class="sec-title"><span class="accent">{emoji}</span>{title}</div>',
                unsafe_allow_html=True)


def kpi(icon, label, value, sub="", color="bleu"):
    st.markdown(f"""
    <div class="kpi-card {color}">
        <span class="icon">{icon}</span>
        <div class="lbl">{label}</div>
        <div class="val">{value}</div>
        <div class="sub">{sub}</div>
    </div>""", unsafe_allow_html=True)


def fmt(v, suffix="FCFA"):
    if v >= 1_000_000_000: return f"{v/1_000_000_000:.2f} Md {suffix}"
    if v >= 1_000_000:     return f"{v/1_000_000:.2f} M {suffix}"
    if v >= 1_000:         return f"{v/1_000:.1f} K {suffix}"
    return f"{v:,.0f} {suffix}"


def chart_header(title, sub=""):
    st.markdown(f'<div class="ctitle">{title}</div><div class="csub">{sub}</div>', unsafe_allow_html=True)


def progress_bar(label, value, max_val, color=XL_BLUE):
    pct = min(100, value / max_val * 100) if max_val else 0
    st.markdown(f"""
    <div class="prog-wrap">
        <div class="prog-labels"><span>{label}</span><span><b>{value:,.0f}</b> FCFA</span></div>
        <div class="prog-bar">
            <div class="prog-fill" style="width:{pct:.1f}%;background:{color};"></div>
        </div>
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# GÉNÉRATION DES DONNÉES
# ══════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def generer_donnees():
    import random as rnd
    rnd.seed(42)
    np.random.seed(42)

    PRODUITS = [
        {"ref":"SAC-SM-001","designation":"Sac Supermarché 30L Standard",     "categorie":"Sacs Supermarché",       "unite":"Paquet 100u","prix_achat":1800,"prix_vente":2500,"seuil":200},
        {"ref":"SAC-SM-002","designation":"Sac Supermarché 50L Renforcé",      "categorie":"Sacs Supermarché",       "unite":"Paquet 100u","prix_achat":2600,"prix_vente":3500,"seuil":200},
        {"ref":"SAC-SM-003","designation":"Sac Supermarché 20L Mini",          "categorie":"Sacs Supermarché",       "unite":"Paquet 200u","prix_achat":1400,"prix_vente":2000,"seuil":300},
        {"ref":"SAC-SM-004","designation":"Sac Supermarché HD Noir 40L",       "categorie":"Sacs Supermarché",       "unite":"Paquet 100u","prix_achat":2200,"prix_vente":3000,"seuil":150},
        {"ref":"SAC-SM-005","designation":"Sac Caisse Transparent 25L",        "categorie":"Sacs Supermarché",       "unite":"Rouleau 500u","prix_achat":4500,"prix_vente":6200,"seuil":100},
        {"ref":"SAC-PB-001","designation":"Sac Poubelle 30L Vert",             "categorie":"Sacs Poubelles",         "unite":"Paquet 20u","prix_achat":650,"prix_vente":950,"seuil":500},
        {"ref":"SAC-PB-002","designation":"Sac Poubelle 50L Noir Renforcé",    "categorie":"Sacs Poubelles",         "unite":"Paquet 20u","prix_achat":850,"prix_vente":1200,"seuil":500},
        {"ref":"SAC-PB-003","designation":"Sac Poubelle 110L Industrial",      "categorie":"Sacs Poubelles",         "unite":"Paquet 10u","prix_achat":1200,"prix_vente":1700,"seuil":300},
        {"ref":"SAC-PB-004","designation":"Sac Poubelle 20L Bleu Ménager",     "categorie":"Sacs Poubelles",         "unite":"Paquet 30u","prix_achat":550,"prix_vente":800,"seuil":400},
        {"ref":"SAC-PB-005","designation":"Sac Poubelle 240L Collectivité",    "categorie":"Sacs Poubelles",         "unite":"Paquet 5u","prix_achat":2200,"prix_vente":3000,"seuil":100},
        {"ref":"GOB-001",   "designation":"Gobelet Plastique 20cl Standard",   "categorie":"Gobelets",               "unite":"Carton 1000u","prix_achat":3500,"prix_vente":5000,"seuil":150},
        {"ref":"GOB-002",   "designation":"Gobelet Plastique 33cl Grand",      "categorie":"Gobelets",               "unite":"Carton 500u","prix_achat":2800,"prix_vente":4000,"seuil":150},
        {"ref":"GOB-003",   "designation":"Gobelet Cristal Transparent 25cl",  "categorie":"Gobelets",               "unite":"Carton 500u","prix_achat":3200,"prix_vente":4500,"seuil":100},
        {"ref":"GOB-004",   "designation":"Gobelet Solo Cup 50cl",             "categorie":"Gobelets",               "unite":"Carton 200u","prix_achat":2400,"prix_vente":3500,"seuil":80},
        {"ref":"GOB-005",   "designation":"Gobelet Dégustation 10cl",          "categorie":"Gobelets",               "unite":"Carton 2000u","prix_achat":3800,"prix_vente":5500,"seuil":120},
        {"ref":"EMB-AL-001","designation":"Barquette Alimentaire PP 500ml",    "categorie":"Emballages Alimentaires","unite":"Carton 400u","prix_achat":4200,"prix_vente":6000,"seuil":100},
        {"ref":"EMB-AL-002","designation":"Barquette Alimentaire PP 1L",       "categorie":"Emballages Alimentaires","unite":"Carton 200u","prix_achat":3600,"prix_vente":5200,"seuil":80},
        {"ref":"EMB-AL-003","designation":"Film Étirable Alimentaire 300m",    "categorie":"Emballages Alimentaires","unite":"Rouleau","prix_achat":5500,"prix_vente":8000,"seuil":50},
        {"ref":"EMB-AL-004","designation":"Sachet Zip PP 1L",                  "categorie":"Emballages Alimentaires","unite":"Paquet 100u","prix_achat":1800,"prix_vente":2600,"seuil":200},
        {"ref":"EMB-IN-001","designation":"Big Bag Industriel 1T FIBC",        "categorie":"Emballages Industriels", "unite":"Unité","prix_achat":8500,"prix_vente":12000,"seuil":30},
        {"ref":"EMB-IN-002","designation":"Fût Plastique HDPE 220L",           "categorie":"Emballages Industriels", "unite":"Unité","prix_achat":12000,"prix_vente":17000,"seuil":20},
        {"ref":"EMB-IN-003","designation":"Jerrycan HDPE 10L",                 "categorie":"Emballages Industriels", "unite":"Unité","prix_achat":2800,"prix_vente":4000,"seuil":60},
        {"ref":"EMB-IN-004","designation":"Bidon HDPE 5L avec bouchon",        "categorie":"Emballages Industriels", "unite":"Unité","prix_achat":1600,"prix_vente":2400,"seuil":80},
        {"ref":"EMB-IN-005","designation":"Sac Industriel PE 50Kg",            "categorie":"Emballages Industriels", "unite":"Paquet 50u","prix_achat":6500,"prix_vente":9000,"seuil":40},
        {"ref":"MAT-PE-001","designation":"Granulés PE-HD (Polyéthylène HD)",  "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":14000,"prix_vente":None,"seuil":100},
        {"ref":"MAT-PP-001","designation":"Granulés PP (Polypropylène)",       "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":15500,"prix_vente":None,"seuil":80},
        {"ref":"MAT-PVC-001","designation":"Granulés PVC Souple",              "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":18000,"prix_vente":None,"seuil":60},
        {"ref":"MAT-REC-001","designation":"Plastique Recyclé Broyé Mix",      "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":7500,"prix_vente":None,"seuil":50},
        {"ref":"MAT-COL-001","designation":"Masterbatch Noir Colorant",        "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":22000,"prix_vente":None,"seuil":40},
        {"ref":"MAT-COL-002","designation":"Masterbatch Vert Colorant",        "categorie":"Matières Premières",     "unite":"Sac 25Kg","prix_achat":23000,"prix_vente":None,"seuil":35},
    ]

    CLIENTS = [
        {"nom":"SOCOCÉ Abidjan Plateau",       "segment":"Grande Distribution",  "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"Carrefour Market Cocody",       "segment":"Grande Distribution",  "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"PROSUMA SA",                    "segment":"Grande Distribution",  "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"Leader Price Yopougon",         "segment":"Grande Distribution",  "ville":"Abidjan",    "zone":"Zone Nord"},
        {"nom":"CDCI Supermarché",              "segment":"Grande Distribution",  "ville":"Abidjan",    "zone":"Zone Nord"},
        {"nom":"NESTLE CI",                     "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"SIC CACAOS",                    "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"SOLIBRA Brasserie",             "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Nord"},
        {"nom":"GONFREVILLE Textile",           "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Nord"},
        {"nom":"SIFCA Groupe",                  "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"PALMCI",                        "segment":"Industriel",           "ville":"San-Pédro",  "zone":"Zone Ouest"},
        {"nom":"CIE Côte d'Ivoire",             "segment":"Industriel",           "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"Commerce Général Yop",          "segment":"Distributeur Local",   "ville":"Abidjan",    "zone":"Zone Nord"},
        {"nom":"DIOULA NÉGOCE",                 "segment":"Distributeur Local",   "ville":"Bouaké",     "zone":"Zone Centre"},
        {"nom":"Maquis & Resto Supplies",       "segment":"Distributeur Local",   "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"AGRO DIST Korhogo",             "segment":"Distributeur Local",   "ville":"Korhogo",    "zone":"Zone Nord"},
        {"nom":"TRANSIT SAN-PEDRO",             "segment":"Distributeur Local",   "ville":"San-Pédro",  "zone":"Zone Ouest"},
        {"nom":"COMMERCE ABENGOUROU",           "segment":"Distributeur Local",   "ville":"Abengourou", "zone":"Zone Est"},
        {"nom":"DIST EXPRESS BOUAKÉ",           "segment":"Distributeur Local",   "ville":"Bouaké",     "zone":"Zone Centre"},
        {"nom":"SOFITEL Abidjan",               "segment":"Hôtellerie & Resto",   "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"Hôtel Ivoire Intercontinental", "segment":"Hôtellerie & Resto",   "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"Groupe LAICO",                  "segment":"Hôtellerie & Resto",   "ville":"Abidjan",    "zone":"Zone Sud"},
        {"nom":"WEST AFRICA TRADING CO.",       "segment":"Export",               "ville":"Dakar",      "zone":"Export"},
        {"nom":"GHANA IMPORT GROUP",            "segment":"Export",               "ville":"Accra",      "zone":"Export"},
        {"nom":"MALI PACKAGING SARL",           "segment":"Export",               "ville":"Bamako",     "zone":"Export"},
    ]

    MODES  = ["Virement bancaire","Chèque certifié","Espèce","Carte bancaire","Traite"]
    POIDS  = [0.40, 0.25, 0.18, 0.10, 0.07]
    SAISON = [0.75,0.70,0.80,0.85,0.90,0.97,0.88,1.12,1.22,1.30,1.45,1.38]
    SEG_W  = {"Grande Distribution":3.2,"Industriel":2.8,"Distributeur Local":1.5,"Hôtellerie & Resto":1.0,"Export":2.2}
    P_PAY  = {"Grande Distribution":0.88,"Industriel":0.85,"Distributeur Local":0.72,"Hôtellerie & Resto":0.91,"Export":0.78}

    from datetime import timedelta
    start = datetime(2022, 1, 1)
    pv = [p for p in PRODUITS if p["prix_vente"]]
    factures = []
    for i in range(850):
        cl  = rnd.choice(CLIENTS)
        d   = start + timedelta(days=rnd.randint(0, 1094))
        sw  = SEG_W.get(cl["segment"], 1.5)
        co  = SAISON[d.month-1] * sw
        nb  = rnd.choices([1,2,3,4], weights=[0.28,0.38,0.24,0.10])[0]
        sel = rnd.sample(pv, min(nb, len(pv)))
        mht = round(sum(p["prix_vente"] * rnd.randint(1, int(10*co)) for p in sel), 0)
        tva = round(mht * 0.18, 0)
        ttc = mht + tva
        ok  = rnd.random() < P_PAY.get(cl["segment"], 0.80)
        mode = rnd.choices(MODES, weights=POIDS)[0] if ok else ""
        factures.append({
            "N° FACTURE": f"F{d.year}-{str(d.month).zfill(2)}-{str(i+1).zfill(4)}",
            "DATE": d, "ANNEE": d.year, "MOIS": d.month,
            "MOIS_NOM": d.strftime("%b %Y"),
            "CLIENT": cl["nom"], "SEGMENT": cl["segment"],
            "VILLE": cl["ville"], "ZONE": cl["zone"],
            "MONTANT HT": mht, "TVA": tva, "MONTANT TTC": ttc,
            "ETAT DE PAIEMENT": "Payée" if ok else "Impayée",
            "MODE DE PAIEMENT": mode,
        })
    df_f = pd.DataFrame(factures).sort_values("DATE").reset_index(drop=True)

    # Stock
    inv = []
    for p in PRODUITS:
        si = rnd.randint(30, 400)
        inv.append({**p, "Stock initial": si, "Entrées": 0, "Sorties": 0, "Stock final": si})
    df_inv = pd.DataFrame(inv)

    ent_rows, sor_rows = [], []
    for p in PRODUITS:
        for _ in range(rnd.randint(12, 45)):
            d   = start + timedelta(days=rnd.randint(0, 1094))
            qte = rnd.randint(10, 600)
            cout= round(p["prix_achat"] * rnd.uniform(0.91, 1.06), 0)
            ent_rows.append({"Date":d,"Référence":p["ref"],"Désignation":p["designation"],
                             "Catégorie":p["categorie"],"Coût d'achat":cout,
                             "Quantité":qte,"Total":round(cout*qte,0)})
        if p["prix_vente"]:
            for _ in range(rnd.randint(8, 40)):
                d   = start + timedelta(days=rnd.randint(0, 1094))
                qte = rnd.randint(5, 350)
                pv_ = round(p["prix_vente"] * rnd.uniform(0.97, 1.10), 0)
                sor_rows.append({"Date":d,"Référence":p["ref"],"Désignation":p["designation"],
                                 "Catégorie":p["categorie"],"Prix de vente":pv_,
                                 "Quantité":qte,"Total":round(pv_*qte,0)})
    df_ent = pd.DataFrame(ent_rows).sort_values("Date").reset_index(drop=True)
    df_sor = pd.DataFrame(sor_rows).sort_values("Date").reset_index(drop=True)

    for i, row in df_inv.iterrows():
        ref = row["ref"]
        e = df_ent[df_ent["Référence"]==ref]["Quantité"].sum()
        s = df_sor[df_sor["Référence"]==ref]["Quantité"].sum() if ref in df_sor["Référence"].values else 0
        sf= max(0, int(row["Stock initial"]) + int(e) - int(s))
        df_inv.at[i,"Entrées"]    = int(e)
        df_inv.at[i,"Sorties"]    = int(s)
        df_inv.at[i,"Stock final"]= sf
        df_inv.at[i,"Valeur"]     = round(sf * row["prix_achat"], 0)
        df_inv.at[i,"Statut"]     = ("Non disponible" if sf==0 else
                                     "Stock faible" if sf<row["seuil"] else "Stock normal")

    # Production
    machines = ["Presse 1 – Injection","Presse 2 – Soufflage","Extrudeuse A","Extrudeuse B","Thermoformeuse"]
    prod_rows = []
    for m in range(36):
        d = start + pd.DateOffset(months=m)
        co = SAISON[d.month-1]
        for mach in machines:
            r  = rnd.uniform(0.73, 0.97)
            pp = int(rnd.randint(50000,140000)*co)
            pr = int(pp*r)
            reb= int(pr*rnd.uniform(0.01,0.055))
            prod_rows.append({
                "Mois":d.strftime("%Y-%m"),"Mois_Label":d.strftime("%b %Y"),
                "Année":d.year,"Machine":mach,
                "Production planifiée":pp,"Production réelle":pr,"Rebuts":reb,
                "Taux rendement":round(r*100,1),
                "Taux rebut":round(reb/pr*100,2) if pr>0 else 0,
            })
    df_prod = pd.DataFrame(prod_rows)
    return df_f, df_inv, df_ent, df_sor, df_prod


# ══════════════════════════════════════════════
# CHARGEMENT
# ══════════════════════════════════════════════
with st.spinner("Chargement du tableau de bord…"):
    df_fact, df_inv, df_ent, df_sor, df_prod = generer_donnees()

# ══════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:20px 0 12px;">
        <div style="font-size:2.2rem;">📦</div>
        <div style="font-size:1.05rem;font-weight:800;color:white;letter-spacing:-0.01em;">MULTIPACK SA</div>
        <div style="font-size:0.71rem;color:#94A3B8;margin-top:2px;">Zone Ind. Yopougon · Abidjan</div>
    </div>
    <hr style="border-color:rgba(255,255,255,0.1);margin:8px 0 16px;">
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", [
        "🏠  Vue d'ensemble",
        "📊  Performance commerciale",
        "🗂️  Analyse par produits",
        "👥  Portefeuille clients",
        "📦  Gestion des stocks",
        "🏭  Production & Rendement",
        "💳  Paiements & Trésorerie",
        "⚠️  Alertes & Suivi PDG",
    ], label_visibility="collapsed")

    st.markdown(f"<hr style='border-color:rgba(255,255,255,0.1);margin:12px 0;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:0.7rem;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px;padding-left:4px;'>Filtres</div>", unsafe_allow_html=True)

    annees = sorted(df_fact["ANNEE"].unique().tolist())
    sel_annees = st.multiselect("Années", annees, default=annees)
    sel_seg = st.multiselect("Segments", df_fact["SEGMENT"].unique().tolist(),
                              default=df_fact["SEGMENT"].unique().tolist())

    st.markdown(f"""
    <div style="margin-top:24px;font-size:0.69rem;color:#475569;text-align:center;line-height:1.6;">
        Mise à jour<br>{datetime.now().strftime('%d/%m/%Y %H:%M')}<br>
        © MULTIPACK SA 2024
    </div>""", unsafe_allow_html=True)

# ── Filtre global
df_f = df_fact[df_fact["ANNEE"].isin(sel_annees) & df_fact["SEGMENT"].isin(sel_seg)]

# ══════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════
now = datetime.now()
st.markdown(f"""
<div class="mp-header">
  <div>
    <h1>📦 Tableau de Bord — Direction Générale</h1>
    <p class="sub">Pilotage Commercial &amp; Reporting · Fabrication de plastiques &amp; emballages</p>
  </div>
  <div class="date-badge">
    <span class="day">{now.strftime('%d')}</span>
    {now.strftime('%b %Y')}<br>
    <span style="opacity:0.65;font-size:0.69rem;">{now.strftime('%H:%M')}</span>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# PAGE : VUE D'ENSEMBLE
# ══════════════════════════════════════════════════════════════
if page == "🏠  Vue d'ensemble":

    ca_total  = df_f["MONTANT TTC"].sum()
    ca_paye   = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    ca_impaye = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]["MONTANT TTC"].sum()
    taux_rec  = ca_paye/ca_total*100 if ca_total else 0
    nb_cl     = df_f["CLIENT"].nunique()
    val_stock = df_inv["Valeur"].sum()
    nb_rupt   = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib   = (df_inv["Statut"]=="Stock faible").sum()
    rend_moy  = df_prod["Taux rendement"].mean()

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("💰","Chiffre d'Affaires Total", fmt(ca_total), f"{len(df_f)} factures","bleu")
    with c2: kpi("✅","Montant Recouvré", fmt(ca_paye), f"Taux : {taux_rec:.1f}%","vert")
    with c3: kpi("⏳","Créances Impayées", fmt(ca_impaye), f"{len(df_f[df_f['ETAT DE PAIEMENT']=='Impayée'])} factures","rouge")
    with c4: kpi("👥","Clients Actifs", str(nb_cl), f"{df_f['SEGMENT'].nunique()} segments","purple")

    st.markdown("<br>", unsafe_allow_html=True)
    c5,c6,c7,c8 = st.columns(4)
    with c5: kpi("🏭","Rendement Production", f"{rend_moy:.1f}%","Toutes machines","bleu")
    with c6: kpi("📦","Valeur du Stock", fmt(val_stock), f"{len(df_inv)} références","purple")
    with c7: kpi("🟡","Stocks Faibles", str(nb_faib),"Sous seuil d'alerte","or")
    with c8: kpi("🔴","Ruptures de Stock", str(nb_rupt),"Réapprovisionnement urgent","rouge" if nb_rupt>0 else "vert")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── GRAPHIQUE 1 : Area chart CA mensuel (comme Excel — AreaChart + LineChart)
    col1, col2 = st.columns([2, 1])
    with col1:
        section("📈","Évolution du CA mensuel (Payé vs Total)")
        ca_m = df_f.copy()
        ca_m["Période"] = ca_m["DATE"].dt.to_period("M").astype(str)
        ca_mois = ca_m.groupby("Période").agg(
            CA_Total=("MONTANT TTC","sum"),
            CA_Paye=("MONTANT TTC", lambda x: x[df_f.loc[x.index,"ETAT DE PAIEMENT"]=="Payée"].sum()),
        ).reset_index()

        # AREA CHART + LINE — exactement comme Excel (AreaChart avec lineChart superposé)
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=ca_mois["Période"], y=ca_mois["CA_Total"],
            name="CA Total", mode="lines",
            fill="tozeroy",
            line=dict(color=XL_BLUE_LIGHT, width=2),
            fillcolor=f"rgba(197,203,251,0.45)",   # #C5CBFB comme Excel
        ))
        fig.add_trace(go.Scatter(
            x=ca_mois["Période"], y=ca_mois["CA_Paye"],
            name="CA Payé", mode="lines",
            fill="tozeroy",
            line=dict(color=XL_BLUE, width=2.5),
            fillcolor=f"rgba(97,114,243,0.35)",    # #6172F3 comme Excel
        ))
        fig = apply_excel_style(fig, height=300)
        fig.update_xaxes(tickangle=-40, tickfont=dict(size=8))
        fig.update_yaxes(tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # ── GRAPHIQUE 2 : Donut chart Payé/Impayé (comme Excel — DoughnutChart, hole=75%)
        section("🥧","Payé vs Impayé")
        stat_d = df_f.groupby("ETAT DE PAIEMENT")["MONTANT TTC"].sum().reset_index()
        fig2 = go.Figure(go.Pie(
            labels=stat_d["ETAT DE PAIEMENT"],
            values=stat_d["MONTANT TTC"],
            hole=0.70,   # Exactement 75% comme Excel (DoughnutChart holeSize=75)
            marker=dict(colors=[XL_BLUE, XL_BLUE_LIGHT],   # couleurs #6172F3 et #C5CBFB comme Excel
                        line=dict(color="white", width=2)),
            textfont=dict(color=FONT_COLOR, size=11),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA<br>%{percent}<extra></extra>",
        ))
        fig2.update_layout(
            paper_bgcolor=CHART_BG, plot_bgcolor=CHART_PLOT_BG,
            height=300, margin=dict(t=10,b=10,l=10,r=10),
            legend=dict(orientation="h", yanchor="bottom", y=-0.12, xanchor="center", x=0.5,
                        font=dict(color=FONT_COLOR, size=10)),
            annotations=[dict(
                text=f"<b>{taux_rec:.0f}%</b><br><span style='font-size:9px'>Payé</span>",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color=FONT_COLOR),
                xanchor="center", yanchor="middle",
            )],
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── ROW 2
    col3, col4, col5 = st.columns([1.2, 1.2, 1.6])

    with col3:
        # DONUT MODE DE PAIEMENT (comme Excel — 2ème DoughnutChart avec rose+bleu)
        section("💳","Modes de Paiement")
        df_paye = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]
        mode_d = df_paye.groupby("MODE DE PAIEMENT")["MONTANT TTC"].sum().reset_index()
        # Couleurs pink+blue comme le donut 3 de l'Excel (#FF6BA7, #6172F3)
        mode_colors = [XL_PINK, XL_BLUE, XL_BLUE_LIGHT, XL_BLUE_DARK, XL_PURPLE]
        fig3 = go.Figure(go.Pie(
            labels=mode_d["MODE DE PAIEMENT"],
            values=mode_d["MONTANT TTC"],
            hole=0.70,
            marker=dict(colors=mode_colors[:len(mode_d)],
                        line=dict(color="white", width=2)),
            textfont=dict(size=10, color=FONT_COLOR),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA (%{percent})<extra></extra>",
        ))
        fig3.update_layout(
            paper_bgcolor=CHART_BG, plot_bgcolor=CHART_PLOT_BG,
            height=260, margin=dict(t=10,b=10,l=5,r=5),
            legend=dict(orientation="v", x=1.05, y=0.5,
                        font=dict(color=FONT_COLOR, size=9)),
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Stock par catégorie — PieChart comme Excel
        section("📦","Stock par Catégorie")
        cat_s = df_inv.groupby("categorie")["Stock final"].sum().reset_index()
        fig4 = go.Figure(go.Pie(
            labels=cat_s["categorie"],
            values=cat_s["Stock final"],
            hole=0,   # PieChart (pas de trou) comme Excel
            marker=dict(colors=EXCEL_PIE_COLORS[:len(cat_s)],
                        line=dict(color="white", width=2)),
            textfont=dict(size=9, color="white"),
            hovertemplate="%{label}<br>%{value:,} unités (%{percent})<extra></extra>",
        ))
        fig4.update_layout(
            paper_bgcolor=CHART_BG, plot_bgcolor=CHART_PLOT_BG,
            height=260, margin=dict(t=10,b=10,l=5,r=5),
            legend=dict(orientation="v", x=1.02, y=0.5,
                        font=dict(color=FONT_COLOR, size=9)),
        )
        st.plotly_chart(fig4, use_container_width=True)

    with col5:
        # BAR CHART Top clients — comme Excel (BarChart horizontal, couleur #6172F3)
        section("🏆","Top 5 Clients — CA Total")
        top5 = df_f.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=True).tail(5).reset_index()
        fig5 = go.Figure(go.Bar(
            x=top5["MONTANT TTC"], y=top5["CLIENT"],
            orientation="h",
            marker_color=XL_BLUE,   # #6172F3 comme Excel BarChart
            hovertemplate="%{y}<br>%{x:,.0f} FCFA<extra></extra>",
            text=top5["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
            textposition="outside",
            textfont=dict(color=FONT_COLOR, size=9),
        ))
        fig5 = apply_excel_style(fig5, height=260, show_legend=False)
        fig5.update_xaxes(tickformat=",.0f", showgrid=True)
        st.plotly_chart(fig5, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE : PERFORMANCE COMMERCIALE
# ══════════════════════════════════════════════════════════════
elif page == "📊  Performance commerciale":
    section("📊","Indicateurs de Performance Commerciale")

    ca = df_f["MONTANT TTC"].sum()
    paye = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    panier = df_f["MONTANT TTC"].mean()
    taux = paye/ca*100 if ca else 0
    ca21 = df_fact[df_fact["ANNEE"]==2022]["MONTANT TTC"].sum()
    ca22 = df_fact[df_fact["ANNEE"]==2023]["MONTANT TTC"].sum()
    crois = (ca22-ca21)/ca21*100 if ca21 else 0

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","CA Total", fmt(ca), f"{len(df_f)} factures","bleu")
    with c2: kpi("✅","Recouvré", fmt(paye), f"{taux:.1f}%","vert")
    with c3: kpi("🛒","Panier Moyen", fmt(panier),"Par facture","purple")
    with c4: kpi("👥","Clients", str(df_f["CLIENT"].nunique()),"Actifs","or")
    with c5: kpi("📈","Croissance 22→23", f"{crois:+.1f}%","Évolution annuelle","vert" if crois>0 else "rouge")

    col1, col2 = st.columns(2)

    with col1:
        # AREA CHART CA mensuel vs Objectif — comme Excel AreaChart+LineChart
        section("📅","CA Mensuel vs Objectif (Area Chart)")
        ca_m2 = df_f.copy()
        ca_m2["Période"] = ca_m2["DATE"].dt.to_period("M").astype(str)
        ca_mois2 = ca_m2.groupby("Période")["MONTANT TTC"].sum().reset_index()
        ca_mois2.columns = ["Période","CA"]
        ca_mois2["Objectif"] = ca_mois2["CA"].mean() * 1.10

        fig = go.Figure()
        # Area (fond coloré) — style Excel AreaChart
        fig.add_trace(go.Scatter(
            x=ca_mois2["Période"], y=ca_mois2["CA"],
            name="CA Réel", mode="lines",
            fill="tozeroy",
            line=dict(color=XL_BLUE, width=2.5),
            fillcolor="rgba(97,114,243,0.20)",
        ))
        # Ligne objectif — style Excel LineChart superposé
        fig.add_trace(go.Scatter(
            x=ca_mois2["Période"], y=ca_mois2["Objectif"],
            name="Objectif +10%", mode="lines",
            line=dict(color=XL_PINK, width=2, dash="dot"),
        ))
        fig = apply_excel_style(fig, height=340)
        fig.update_xaxes(tickangle=-40, tickfont=dict(size=8))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # BAR CHART empilé par statut — comme Excel BarChart couleur #6172F3
        section("📊","CA par Statut de Paiement (Bar Chart)")
        df_f2 = df_f.copy()
        df_f2["Période"] = df_f2["DATE"].dt.to_period("M").astype(str)
        ca_stat = df_f2.groupby(["Période","ETAT DE PAIEMENT"])["MONTANT TTC"].sum().reset_index()
        fig2 = go.Figure()
        for statut, color in [("Payée", XL_BLUE), ("Impayée", XL_BLUE_LIGHT)]:
            d = ca_stat[ca_stat["ETAT DE PAIEMENT"]==statut]
            fig2.add_trace(go.Bar(
                x=d["Période"], y=d["MONTANT TTC"], name=statut,
                marker_color=color,
            ))
        fig2.update_layout(barmode="stack")
        fig2 = apply_excel_style(fig2, height=340)
        fig2.update_xaxes(tickangle=-40, tickfont=dict(size=8))
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        # BAR HORIZONTAL top clients — couleur #6172F3 comme Excel
        section("🏆","Top 10 Clients par CA (Bar Chart)")
        top10 = df_f.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=True).tail(10).reset_index()
        fig3 = go.Figure(go.Bar(
            x=top10["MONTANT TTC"], y=top10["CLIENT"],
            orientation="h",
            marker_color=XL_BLUE,
            hovertemplate="%{y}: %{x:,.0f} FCFA<extra></extra>",
        ))
        fig3 = apply_excel_style(fig3, height=360, show_legend=False)
        fig3.update_xaxes(tickformat=",.0f")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # AREA CHART comparatif annuel — comme Excel (séries multiples)
        section("📈","Comparatif CA par Année (Area Chart)")
        ca_ann = df_f.groupby(["ANNEE","MOIS"])["MONTANT TTC"].sum().reset_index()
        mois_l = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jun",
                  7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
        ca_ann["Mois_L"] = ca_ann["MOIS"].map(mois_l)
        colors_ann = [XL_BLUE, XL_PINK, XL_PURPLE]
        fig4 = go.Figure()
        for j, an in enumerate(sorted(ca_ann["ANNEE"].unique())):
            d = ca_ann[ca_ann["ANNEE"]==an]
            fig4.add_trace(go.Scatter(
                x=d["Mois_L"], y=d["MONTANT TTC"],
                name=str(an), mode="lines+markers",
                fill="tozeroy",
                line=dict(color=colors_ann[j % len(colors_ann)], width=2),
                fillcolor=f"rgba({int(colors_ann[j % len(colors_ann)][1:3],16)},"
                          f"{int(colors_ann[j % len(colors_ann)][3:5],16)},"
                          f"{int(colors_ann[j % len(colors_ann)][5:7],16)},0.12)",
                marker=dict(size=4),
            ))
        fig4 = apply_excel_style(fig4, height=360)
        st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE : ANALYSE PAR PRODUITS
# ══════════════════════════════════════════════════════════════
elif page == "🗂️  Analyse par produits":
    section("🗂️","Performance par Famille de Produits")

    sor_agg = df_sor.groupby("Catégorie").agg(
        CA_Ventes=("Total","sum"), Nb_Trans=("Total","count"), Qte=("Quantité","sum")).reset_index()
    ent_agg = df_ent.groupby("Catégorie").agg(
        Achats=("Total","sum"), Qte_A=("Quantité","sum")).reset_index()
    perf = sor_agg.merge(ent_agg, on="Catégorie", how="left").fillna(0)
    perf["Marge"] = perf["CA_Ventes"] - perf["Achats"]
    perf["Taux%"] = np.where(perf["CA_Ventes"]>0, perf["Marge"]/perf["CA_Ventes"]*100, 0)

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("📦","CA Ventes Produits", fmt(sor_agg["CA_Ventes"].sum()),"Total catégories","bleu")
    with c2: kpi("🛒","Coût Achats Matières", fmt(ent_agg["Achats"].sum()),"Coût d'achat","or")
    with c3: kpi("📊","Marge Brute Estimée", fmt(perf["Marge"].sum()),"CA – Achats","vert")
    with c4:
        tm = perf["Marge"].sum()/sor_agg["CA_Ventes"].sum()*100 if sor_agg["CA_Ventes"].sum()>0 else 0
        kpi("💹","Taux de Marge Moyen", f"{tm:.1f}%","Sur toutes catégories","purple")

    col1, col2 = st.columns(2)

    with col1:
        # BAR CHART groupé — couleur #6172F3 principale, #C5CBFB secondaire (comme Excel)
        section("📊","CA Ventes vs Achats par Catégorie (Bar Chart)")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=perf["Catégorie"], y=perf["CA_Ventes"],
                              name="CA Ventes", marker_color=XL_BLUE))
        fig.add_trace(go.Bar(x=perf["Catégorie"], y=perf["Achats"],
                              name="Coût Achats", marker_color=XL_BLUE_LIGHT))
        fig.update_layout(barmode="group")
        fig = apply_excel_style(fig, height=320)
        fig.update_xaxes(tickangle=-25, tickfont=dict(size=8))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # PIE CHART CA par catégorie — couleurs Excel pie standard
        section("🥧","Répartition CA par Catégorie (Pie Chart)")
        sor_cat = df_sor.groupby("Catégorie")["Total"].sum().reset_index()
        fig2 = go.Figure(go.Pie(
            labels=sor_cat["Catégorie"],
            values=sor_cat["Total"],
            hole=0,   # PieChart comme Excel
            marker=dict(colors=EXCEL_PIE_COLORS[:len(sor_cat)],
                        line=dict(color="white", width=2)),
            textfont=dict(size=9),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA (%{percent})<extra></extra>",
        ))
        fig2.update_layout(
            paper_bgcolor=CHART_BG, height=320,
            margin=dict(t=10,b=30,l=5,r=5),
            legend=dict(orientation="h", yanchor="top", y=-0.05,
                        xanchor="center", x=0.5, font=dict(color=FONT_COLOR, size=9)),
        )
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        # AREA CHART ventes mensuelles par catégorie — style #6172F3 et #51459E (Stock Excel)
        section("📉","Évolution Mensuelle Ventes (Area Chart)")
        df_sor_m = df_sor.copy()
        df_sor_m["Mois"] = pd.to_datetime(df_sor_m["Date"]).dt.to_period("M").astype(str)
        sor_mc = df_sor_m.groupby(["Mois","Catégorie"])["Total"].sum().reset_index()
        cats = sor_mc["Catégorie"].unique()
        area_colors = [XL_BLUE, XL_PINK, XL_PURPLE, XL_BLUE_DARK, XL_BLUE_LIGHT, "#F7B731","#20BF6B"]
        fig3 = go.Figure()
        for j, cat in enumerate(cats):
            d = sor_mc[sor_mc["Catégorie"]==cat]
            c = area_colors[j % len(area_colors)]
            fig3.add_trace(go.Scatter(
                x=d["Mois"], y=d["Total"], name=cat,
                mode="lines", fill="tozeroy",
                line=dict(color=c, width=1.8),
                fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.12)",
            ))
        fig3 = apply_excel_style(fig3, height=320)
        fig3.update_xaxes(tickangle=-40, tickfont=dict(size=7))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # BAR HORIZONTAL top produits — #6172F3 comme Excel
        section("🏅","Top 15 Produits (Bar Chart)")
        top_p = df_sor.groupby("Désignation")["Total"].sum().sort_values(ascending=True).tail(12).reset_index()
        fig4 = go.Figure(go.Bar(
            x=top_p["Total"], y=top_p["Désignation"],
            orientation="h",
            marker_color=XL_BLUE,
            hovertemplate="%{y}<br>%{x:,.0f} FCFA<extra></extra>",
        ))
        fig4 = apply_excel_style(fig4, height=320, show_legend=False)
        fig4.update_xaxes(tickformat=",.0f")
        st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE : PORTEFEUILLE CLIENTS
# ══════════════════════════════════════════════════════════════
elif page == "👥  Portefeuille clients":
    section("👥","Analyse Portefeuille Clients")

    cl_stats = df_f.groupby(["CLIENT","SEGMENT","ZONE"]).agg(
        CA_Total=("MONTANT TTC","sum"),
        Nb_Cmd=("MONTANT TTC","count"),
        Panier=("MONTANT TTC","mean"),
        CA_Paye=("MONTANT TTC", lambda x: x[df_f.loc[x.index,"ETAT DE PAIEMENT"]=="Payée"].sum()),
        CA_Imp=("MONTANT TTC",  lambda x: x[df_f.loc[x.index,"ETAT DE PAIEMENT"]=="Impayée"].sum()),
    ).reset_index()
    cl_stats["Taux_R"] = np.where(cl_stats["CA_Total"]>0,
                                   cl_stats["CA_Paye"]/cl_stats["CA_Total"]*100, 0).round(1)
    cl_stats = cl_stats.sort_values("CA_Total", ascending=False)

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("👥","Total Clients", str(len(cl_stats)),"Au moins 1 commande","bleu")
    with c2: kpi("⭐","Clients VIP", str(max(1,int(len(cl_stats)*0.2))),"Top 20% CA","or")
    with c3: kpi("⚠️","Clients Impayés", str((cl_stats["CA_Imp"]>0).sum()),"À relancer","rouge")
    with c4: kpi("🔄","Clients Fidèles", str((cl_stats["Nb_Cmd"]>=5).sum()),"≥5 commandes","vert")

    col1, col2 = st.columns(2)
    with col1:
        # DONUT segmentation — couleurs #6172F3 et #C5CBFB comme Excel
        section("🗂️","CA par Segment (Donut Chart)")
        seg_ca = cl_stats.groupby("SEGMENT")["CA_Total"].sum().reset_index()
        seg_colors = [XL_BLUE, XL_PINK, XL_PURPLE, XL_BLUE_LIGHT, XL_BLUE_DARK]
        fig = go.Figure(go.Pie(
            labels=seg_ca["SEGMENT"], values=seg_ca["CA_Total"],
            hole=0.65,
            marker=dict(colors=seg_colors[:len(seg_ca)], line=dict(color="white",width=2)),
            textfont=dict(size=10, color=FONT_COLOR),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA (%{percent})<extra></extra>",
        ))
        fig.update_layout(paper_bgcolor=CHART_BG, height=320, margin=dict(t=10,b=30,l=5,r=5),
                           legend=dict(orientation="h",yanchor="top",y=-0.05,
                                       xanchor="center",x=0.5,font=dict(color=FONT_COLOR,size=9)))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # BAR CA par zone — couleur #6172F3/#C5CBFB empilé comme Excel
        section("🗺️","CA par Zone Géographique (Bar Chart)")
        zone_ca = cl_stats.groupby("ZONE")[["CA_Paye","CA_Imp"]].sum().reset_index()
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=zone_ca["ZONE"], y=zone_ca["CA_Paye"],
                               name="CA Payé", marker_color=XL_BLUE))
        fig2.add_trace(go.Bar(x=zone_ca["ZONE"], y=zone_ca["CA_Imp"],
                               name="CA Impayé", marker_color=XL_BLUE_LIGHT))
        fig2.update_layout(barmode="stack")
        fig2 = apply_excel_style(fig2, height=320)
        fig2.update_xaxes(tickangle=-20, tickfont=dict(size=9))
        st.plotly_chart(fig2, use_container_width=True)

    # Tableau
    section("📋","Fiche Clients Complète")
    disp = cl_stats.copy()
    for c in ["CA_Total","CA_Paye","CA_Imp","Panier"]:
        disp[c] = disp[c].apply(lambda x: f"{x:,.0f} FCFA")
    disp["Taux_R"] = cl_stats["Taux_R"].apply(lambda x: f"{x:.1f}%")
    disp = disp.rename(columns={"CLIENT":"Client","SEGMENT":"Segment","ZONE":"Zone",
                                  "CA_Total":"CA Total","Nb_Cmd":"Commandes","Panier":"Panier Moy.",
                                  "CA_Paye":"Payé","CA_Imp":"Impayé","Taux_R":"Taux Recouv."})
    st.dataframe(disp, use_container_width=True, height=380)


# ══════════════════════════════════════════════════════════════
# PAGE : GESTION DES STOCKS
# ══════════════════════════════════════════════════════════════
elif page == "📦  Gestion des stocks":
    section("📦","État de l'Inventaire MULTIPACK")

    val_stock = df_inv["Valeur"].sum()
    nb_rupt   = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib   = (df_inv["Statut"]=="Stock faible").sum()
    nb_norm   = (df_inv["Statut"]=="Stock normal").sum()

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","Valeur Stock", fmt(val_stock), f"{len(df_inv)} réf.","bleu")
    with c2: kpi("✅","Stock Normal", str(nb_norm),"Niveau OK","vert")
    with c3: kpi("🟡","Stock Faible", str(nb_faib),"Sous seuil","or")
    with c4: kpi("🔴","Ruptures", str(nb_rupt),"Stock = 0","rouge")
    with c5:
        rot = df_sor["Quantité"].sum()/max(df_inv["Stock final"].sum(),1)
        kpi("🔄","Rotation", f"{rot:.1f}x","Sorties/Stock","purple")

    col1, col2, col3 = st.columns([1.6, 1.4, 1])

    with col1:
        # BAR CHART stock par catégorie — couleur #6172F3 comme Excel
        section("📊","Stock Final par Catégorie (Bar Chart)")
        cat_s = df_inv.groupby("categorie")["Stock final"].sum().reset_index()
        fig = go.Figure(go.Bar(
            x=cat_s["categorie"], y=cat_s["Stock final"],
            marker_color=XL_BLUE,
            text=cat_s["Stock final"].apply(lambda x: f"{x:,}"),
            textposition="outside", textfont=dict(size=9, color=FONT_COLOR),
            hovertemplate="%{x}<br>%{y:,} unités<extra></extra>",
        ))
        fig = apply_excel_style(fig, height=300, show_legend=False)
        fig.update_xaxes(tickangle=-20, tickfont=dict(size=8))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # PIE CHART valeur stock — couleurs Excel standard pie
        section("💰","Valeur Stock par Catégorie (Pie Chart)")
        cat_v = df_inv.groupby("categorie")["Valeur"].sum().reset_index()
        fig2 = go.Figure(go.Pie(
            labels=cat_v["categorie"], values=cat_v["Valeur"],
            hole=0,
            marker=dict(colors=EXCEL_PIE_COLORS[:len(cat_v)],
                        line=dict(color="white", width=2)),
            textfont=dict(size=9),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA (%{percent})<extra></extra>",
        ))
        fig2.update_layout(paper_bgcolor=CHART_BG, height=300,
                            margin=dict(t=10,b=30,l=5,r=5),
                            legend=dict(orientation="h",yanchor="top",y=-0.05,
                                        xanchor="center",x=0.5,font=dict(color=FONT_COLOR,size=8)))
        st.plotly_chart(fig2, use_container_width=True)

    with col3:
        section("🚦","Statuts")
        for stat, color, bg, emoji in [
            ("Stock normal","#276749","#F0FFF4","✅"),
            ("Stock faible","#744210","#FFFFF0","⚠️"),
            ("Non disponible","#742A2A","#FFF5F5","🔴"),
        ]:
            nb  = (df_inv["Statut"]==stat).sum()
            pct = nb/len(df_inv)*100 if len(df_inv)>0 else 0
            st.markdown(f"""
            <div style="background:{bg};border-radius:10px;padding:14px 16px;margin:6px 0;
                        border:1px solid {UI['border']};">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <span style="font-size:0.8rem;color:{color};">{emoji} {stat}</span>
                    <span style="font-size:1.1rem;font-weight:800;color:{color};">{nb}</span>
                </div>
                <div style="height:5px;background:#E2E8F0;border-radius:3px;margin-top:8px;">
                    <div style="height:5px;width:{pct:.0f}%;background:{color};border-radius:3px;"></div>
                </div>
                <div style="font-size:0.69rem;color:{color};opacity:0.8;margin-top:4px;">{pct:.0f}%</div>
            </div>""", unsafe_allow_html=True)

    # AREA CHARTS entrées/sorties — style #51459E comme Gestion_de_stock.xlsm
    col4, col5 = st.columns(2)
    with col4:
        section("📥","Flux Achats Mensuels (Area Chart)")
        df_em = df_ent.copy()
        df_em["Mois"] = pd.to_datetime(df_em["Date"]).dt.to_period("M").astype(str)
        em = df_em.groupby(["Mois","Catégorie"])["Total"].sum().reset_index()
        cats = em["Catégorie"].unique()
        # Couleur #51459E comme le fichier stock Excel
        sto_colors = [XL_PURPLE, XL_BLUE, XL_PINK, XL_BLUE_LIGHT, XL_BLUE_DARK, "#F7B731","#20BF6B"]
        fig4 = go.Figure()
        for j, cat in enumerate(cats):
            d = em[em["Catégorie"]==cat]
            c = sto_colors[j % len(sto_colors)]
            fig4.add_trace(go.Scatter(
                x=d["Mois"], y=d["Total"], name=cat, mode="lines",
                fill="tozeroy",
                line=dict(color=c, width=2),
                fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.15)",
            ))
        fig4 = apply_excel_style(fig4, height=300)
        fig4.update_xaxes(tickangle=-40, tickfont=dict(size=7))
        st.plotly_chart(fig4, use_container_width=True)

    with col5:
        section("📤","Flux Ventes Mensuels (Area Chart)")
        df_sm = df_sor.copy()
        df_sm["Mois"] = pd.to_datetime(df_sm["Date"]).dt.to_period("M").astype(str)
        sm = df_sm.groupby(["Mois","Catégorie"])["Total"].sum().reset_index()
        fig5 = go.Figure()
        for j, cat in enumerate(sm["Catégorie"].unique()):
            d = sm[sm["Catégorie"]==cat]
            c = sto_colors[j % len(sto_colors)]
            fig5.add_trace(go.Scatter(
                x=d["Mois"], y=d["Total"], name=cat, mode="lines",
                fill="tozeroy",
                line=dict(color=c, width=2),
                fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.15)",
            ))
        fig5 = apply_excel_style(fig5, height=300)
        fig5.update_xaxes(tickangle=-40, tickfont=dict(size=7))
        st.plotly_chart(fig5, use_container_width=True)

    section("📋","Inventaire Détaillé")
    inv_d = df_inv[["ref","designation","categorie","unite","seuil","Stock initial",
                     "Entrées","Sorties","Stock final","Valeur","Statut"]].copy()
    inv_d["Valeur"] = inv_d["Valeur"].apply(lambda x: f"{x:,.0f} FCFA")
    inv_d = inv_d.rename(columns={"ref":"Réf","designation":"Désignation",
                                    "categorie":"Catégorie","unite":"Unité","seuil":"Seuil"})
    st.dataframe(inv_d, use_container_width=True, height=380)


# ══════════════════════════════════════════════════════════════
# PAGE : PRODUCTION & RENDEMENT
# ══════════════════════════════════════════════════════════════
elif page == "🏭  Production & Rendement":
    section("🏭","Tableau de Bord Production")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("⚙️","Production Totale", f"{df_prod['Production réelle'].sum():,.0f}", "unités","bleu")
    with c2: kpi("📈","Rendement Moyen", f"{df_prod['Taux rendement'].mean():.1f}%","Toutes machines","vert")
    with c3: kpi("🗑️","Taux Rebut Moyen", f"{df_prod['Taux rebut'].mean():.2f}%","Unités perdues","rouge")
    with c4: kpi("🔧","Total Rebuts", f"{df_prod['Rebuts'].sum():,.0f}","Unités non conformes","or")

    col1, col2 = st.columns(2)
    with col1:
        # BAR + LINE superposés — comme Excel BarChart+LineChart
        section("📅","Production Réelle vs Planifiée (Bar + Line)")
        pm = df_prod.groupby("Mois_Label").agg(
            Planifiée=("Production planifiée","sum"),
            Réelle=("Production réelle","sum"),
            Rebuts=("Rebuts","sum"),
        ).reset_index()
        try:
            pm["sk"] = pd.to_datetime(pm["Mois_Label"], format="%b %Y")
            pm = pm.sort_values("sk")
        except: pass

        fig = go.Figure()
        fig.add_trace(go.Bar(x=pm["Mois_Label"], y=pm["Planifiée"],
                              name="Planifiée", marker_color=XL_BLUE_LIGHT, opacity=0.6))
        fig.add_trace(go.Bar(x=pm["Mois_Label"], y=pm["Réelle"],
                              name="Réelle", marker_color=XL_BLUE, opacity=0.9))
        # Line superposée pour rebuts — comme Excel LineChart sur BarChart
        fig.add_trace(go.Scatter(x=pm["Mois_Label"], y=pm["Rebuts"],
                                  name="Rebuts", mode="lines+markers",
                                  line=dict(color=XL_PINK, width=2),
                                  marker=dict(size=4),
                                  yaxis="y2"))
        fig.update_layout(
            barmode="overlay",
            yaxis2=dict(overlaying="y", side="right", showgrid=False,
                        tickfont=dict(color=XL_PINK, size=9),
                        title=dict(text="Rebuts", font=dict(color=XL_PINK))),
        )
        fig = apply_excel_style(fig, height=340)
        fig.update_xaxes(tickangle=-40, tickfont=dict(size=7))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # BAR CHART rendement par machine — #6172F3 comme Excel BarChart
        section("⚙️","Rendement par Machine (Bar Chart)")
        rm = df_prod.groupby("Machine").agg(
            Rendement=("Taux rendement","mean"),
            Rebut=("Taux rebut","mean"),
        ).reset_index().sort_values("Rendement", ascending=True)
        fig2 = go.Figure(go.Bar(
            x=rm["Rendement"], y=rm["Machine"],
            orientation="h",
            marker_color=[XL_BLUE if v>=85 else XL_BLUE_LIGHT if v>=75 else XL_PINK
                          for v in rm["Rendement"]],
            text=rm["Rendement"].apply(lambda x: f"{x:.1f}%"),
            textposition="outside", textfont=dict(size=10, color=FONT_COLOR),
            hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
        ))
        fig2.add_vline(x=85, line_dash="dash", line_color=XL_PINK, line_width=1.5,
                        annotation_text="Obj. 85%",
                        annotation_font=dict(color=XL_PINK, size=10))
        fig2 = apply_excel_style(fig2, height=340, show_legend=False)
        fig2.update_xaxes(ticksuffix="%", range=[60, 105])
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        # AREA CHART rebuts par machine — #51459E style Excel stock
        section("📉","Évolution Taux Rebut (Area Chart)")
        rb_m = df_prod.groupby(["Mois_Label","Machine"])["Taux rebut"].mean().reset_index()
        try:
            rb_m["sk"] = pd.to_datetime(rb_m["Mois_Label"], format="%b %Y")
            rb_m = rb_m.sort_values("sk")
        except: pass
        fig3 = go.Figure()
        mach_colors = [XL_PURPLE, XL_BLUE, XL_PINK, XL_BLUE_LIGHT, XL_BLUE_DARK]
        for j, mach in enumerate(df_prod["Machine"].unique()):
            d = rb_m[rb_m["Machine"]==mach]
            c = mach_colors[j % len(mach_colors)]
            fig3.add_trace(go.Scatter(
                x=d["Mois_Label"], y=d["Taux rebut"],
                name=mach.split("–")[0].strip(), mode="lines",
                fill="tozeroy",
                line=dict(color=c, width=2),
                fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.12)",
            ))
        fig3 = apply_excel_style(fig3, height=300)
        fig3.update_xaxes(tickangle=-40, tickfont=dict(size=7))
        fig3.update_yaxes(ticksuffix="%")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # BAR CHART production par machine cumulée
        section("🏭","Production Cumulée par Machine (Bar Chart)")
        pm2 = df_prod.groupby("Machine")["Production réelle"].sum().reset_index().sort_values("Production réelle")
        fig4 = go.Figure(go.Bar(
            x=pm2["Production réelle"], y=pm2["Machine"],
            orientation="h", marker_color=XL_BLUE,
            text=pm2["Production réelle"].apply(lambda x: f"{x/1e6:.1f}M"),
            textposition="outside", textfont=dict(size=10, color=FONT_COLOR),
        ))
        fig4 = apply_excel_style(fig4, height=300, show_legend=False)
        fig4.update_xaxes(tickformat=".2s")
        st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE : PAIEMENTS & TRÉSORERIE
# ══════════════════════════════════════════════════════════════
elif page == "💳  Paiements & Trésorerie":
    section("💳","Analyse des Paiements & Trésorerie")

    df_p = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]
    df_i = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]
    taux  = len(df_p)/len(df_f)*100 if len(df_f) else 0
    mode_top = df_p["MODE DE PAIEMENT"].value_counts().idxmax() if len(df_p)>0 else "N/A"

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","Total Recouvré", fmt(df_p["MONTANT TTC"].sum()),"Paiements reçus","vert")
    with c2: kpi("⏳","Total Impayé", fmt(df_i["MONTANT TTC"].sum()), f"{len(df_i)} fact.","rouge")
    with c3: kpi("📊","Taux Recouvrement", f"{taux:.1f}%","En nb factures","bleu")
    with c4: kpi("🧾","Panier Moyen Payé", fmt(df_p["MONTANT TTC"].mean()),"Par facture réglée","or")
    with c5: kpi("🏆","Mode Dominant", mode_top.split()[0],"Le + utilisé","purple")

    col1, col2 = st.columns(2)
    with col1:
        # DONUT mode paiement — #FF6BA7 + #6172F3 comme Excel chart5
        section("🥧","Répartition Modes de Paiement (Donut Chart)")
        md = df_p.groupby("MODE DE PAIEMENT")["MONTANT TTC"].sum().reset_index()
        mode_colors2 = [XL_PINK, XL_BLUE, XL_PURPLE, XL_BLUE_LIGHT, XL_BLUE_DARK]
        fig = go.Figure(go.Pie(
            labels=md["MODE DE PAIEMENT"], values=md["MONTANT TTC"],
            hole=0.65,
            marker=dict(colors=mode_colors2[:len(md)], line=dict(color="white",width=2)),
            textfont=dict(size=10, color=FONT_COLOR),
            hovertemplate="%{label}<br>%{value:,.0f} FCFA (%{percent})<extra></extra>",
        ))
        fig.update_layout(paper_bgcolor=CHART_BG, height=320,
                           margin=dict(t=10,b=30,l=5,r=5),
                           legend=dict(orientation="h",yanchor="top",y=-0.05,
                                       xanchor="center",x=0.5,font=dict(color=FONT_COLOR,size=9)))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # BAR CHART encaissements mensuels par mode — #6172F3/#C5CBFB empilé
        section("📅","Encaissements par Mode & Mois (Bar Chart)")
        df_pm = df_p.copy()
        df_pm["Période"] = df_pm["DATE"].dt.to_period("M").astype(str)
        mm = df_pm.groupby(["Période","MODE DE PAIEMENT"])["MONTANT TTC"].sum().reset_index()
        modes_u = mm["MODE DE PAIEMENT"].unique()
        bar_mode_colors = [XL_BLUE, XL_PINK, XL_PURPLE, XL_BLUE_LIGHT, XL_BLUE_DARK]
        fig2 = go.Figure()
        for j, m in enumerate(modes_u):
            d = mm[mm["MODE DE PAIEMENT"]==m]
            fig2.add_trace(go.Bar(x=d["Période"], y=d["MONTANT TTC"],
                                   name=m, marker_color=bar_mode_colors[j % len(bar_mode_colors)]))
        fig2.update_layout(barmode="stack")
        fig2 = apply_excel_style(fig2, height=320)
        fig2.update_xaxes(tickangle=-40, tickfont=dict(size=8))
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns([1.2, 1.8])
    with col3:
        section("🔴","Top Impayés à Recouvrer")
        top_i = df_i.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=False).head(10)
        mx = top_i.max()
        for cl, v in top_i.items():
            progress_bar(cl[:24], v, mx, XL_PINK)

    with col4:
        # AREA CHART taux recouvrement mensuel — #6172F3 comme Excel
        section("📈","Taux de Recouvrement Mensuel (Area Chart)")
        df_fm = df_f.copy()
        df_fm["Période"] = df_fm["DATE"].dt.to_period("M").astype(str)
        rec_m = df_fm.groupby("Période").apply(
            lambda g: g[g["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()/g["MONTANT TTC"].sum()*100
            if g["MONTANT TTC"].sum()>0 else 0
        ).reset_index()
        rec_m.columns = ["Période","Taux"]
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=rec_m["Période"], y=rec_m["Taux"],
            mode="lines+markers",
            line=dict(color=XL_BLUE, width=2.5),
            fill="tozeroy",
            fillcolor="rgba(97,114,243,0.15)",
            marker=dict(size=4, color=XL_BLUE),
        ))
        fig3.add_hline(y=80, line_dash="dash", line_color=XL_PINK, line_width=1.5,
                        annotation_text="Objectif 80%",
                        annotation_font=dict(color=XL_PINK, size=10))
        fig3 = apply_excel_style(fig3, height=300)
        fig3.update_xaxes(tickangle=-40, tickfont=dict(size=8))
        fig3.update_yaxes(ticksuffix="%", range=[50, 105])
        st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE : ALERTES & SUIVI PDG
# ══════════════════════════════════════════════════════════════
elif page == "⚠️  Alertes & Suivi PDG":

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1E2A4A,#2D3A6B);border-radius:14px;
                padding:20px 26px;margin-bottom:20px;border:1px solid {UI['border']};">
        <div style="font-size:1.1rem;font-weight:800;color:white;">
            📋 Note de Synthèse Exécutive — PDG MULTIPACK SA
        </div>
        <div style="font-size:0.78rem;color:#94A3B8;margin-top:3px;">
            {datetime.now().strftime('%A %d %B %Y · %H:%M')}
        </div>
    </div>""", unsafe_allow_html=True)

    ca_total  = df_f["MONTANT TTC"].sum()
    ca_paye   = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    ca_impaye = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]["MONTANT TTC"].sum()
    taux_rec  = ca_paye/ca_total*100 if ca_total>0 else 0
    nb_rupt   = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib   = (df_inv["Statut"]=="Stock faible").sum()
    rend_moy  = df_prod["Taux rendement"].mean()

    c1,c2,c3 = st.columns(3)
    for col, title, val, st2, ok_thr, danger_thr in [
        (c1,"💰 TRÉSORERIE", f"{taux_rec:.1f}% recouvrement", f"{fmt(ca_impaye)} impayés", 80, 65),
        (c2,"📦 STOCKS", f"{nb_rupt} rupture(s) · {nb_faib} faible(s)", f"{fmt(df_inv['Valeur'].sum())} en stock", 0, 2),
        (c3,"🏭 PRODUCTION", f"{rend_moy:.1f}% rendement moyen", f"{df_prod['Rebuts'].sum():,.0f} rebuts", 85, 75),
    ]:
        v_num = taux_rec if "TRÉSO" in title else (nb_rupt+nb_faib if "STOCK" in title else rend_moy)
        if "STOCK" in title:
            status = ("🟢 BON" if v_num==0 else "🟡 ATTENTION" if v_num<=2 else "🔴 CRITIQUE")
            col_s = (UI["vert"] if v_num==0 else UI["or"] if v_num<=2 else UI["rouge"])
        else:
            status = ("🟢 BON" if v_num>=ok_thr else "🟡 ATTENTION" if v_num>=danger_thr else "🔴 CRITIQUE")
            col_s = (UI["vert"] if v_num>=ok_thr else UI["or"] if v_num>=danger_thr else UI["rouge"])
        col.markdown(f"""
        <div style="background:{UI['card']};border-radius:12px;padding:18px;
                    border:1px solid {UI['border']};border-top:4px solid {col_s};
                    box-shadow:0 2px 8px rgba(0,0,0,0.06);">
            <div style="font-size:0.72rem;font-weight:700;color:{UI['muted']};text-transform:uppercase;
                        letter-spacing:0.07em;">{title}</div>
            <div style="font-size:1.3rem;font-weight:800;color:{UI['text']};margin:8px 0 2px;">{val}</div>
            <div style="font-size:0.76rem;color:{UI['muted']};">{st2}</div>
            <div style="margin-top:10px;font-size:0.82rem;font-weight:600;color:{col_s};">{status}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_al, col_reco = st.columns(2)

    with col_al:
        section("🚨","Alertes Actives")
        rupt = df_inv[df_inv["Statut"]=="Non disponible"]
        faib = df_inv[df_inv["Statut"]=="Stock faible"]
        for _, r in rupt.iterrows():
            st.markdown(f'<div class="alert-r">🔴<div><b>RUPTURE</b> – {r["designation"]}<br><small>{r["categorie"]} · Stock = 0 unité · URGENT</small></div></div>', unsafe_allow_html=True)
        for _, r in faib.head(5).iterrows():
            st.markdown(f'<div class="alert-y">⚠️<div><b>STOCK FAIBLE</b> – {r["designation"]}<br><small>{r["categorie"]} · Restant : {int(r["Stock final"])} u. / Seuil : {int(r["seuil"])} u.</small></div></div>', unsafe_allow_html=True)
        imp = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"].groupby("CLIENT")["MONTANT TTC"].sum()
        for cl, v in imp.nlargest(4).items():
            st.markdown(f'<div class="alert-r">💸<div><b>IMPAYÉ</b> – {cl}<br><small>Créance : {v:,.0f} FCFA · Relance prioritaire</small></div></div>', unsafe_allow_html=True)
        if nb_rupt==0 and nb_faib==0:
            st.markdown('<div class="alert-g">✅ Stocks OK — aucune rupture.</div>', unsafe_allow_html=True)

    with col_reco:
        section("💡","Recommandations Stratégiques")
        recos = []
        if taux_rec < 80:
            recos.append(("r","🔴 Trésorerie", f"Taux de recouvrement {taux_rec:.1f}% < 80%. Lancer une campagne de relance sur {len(df_f[df_f['ETAT DE PAIEMENT']=='Impayée'])} factures impayées ({fmt(ca_impaye)})."))
        if nb_rupt > 0:
            recos.append(("r","🔴 Stock", f"{nb_rupt} produit(s) en rupture totale. Commander en urgence."))
        if nb_faib > 3:
            recos.append(("y","🟡 Stock", f"{nb_faib} références en stock faible. Planifier les réapprovisionnements."))
        if rend_moy < 82:
            recos.append(("y","🟡 Production", f"Rendement moyen {rend_moy:.1f}%. Planifier la maintenance préventive."))
        top_seg = df_f.groupby("SEGMENT")["MONTANT TTC"].sum().idxmax()
        recos.append(("g","🟢 Commercial", f"Segment '{top_seg}' = meilleur contributeur CA. Renforcer la prospection."))
        top_zone = df_f.groupby("ZONE")["MONTANT TTC"].sum().idxmax()
        recos.append(("g","🟢 Géographie", f"Zone '{top_zone}' = zone la plus rentable. Augmenter la couverture."))
        for typ, titre, msg in recos:
            cls = {"r":"alert-r","y":"alert-y","g":"alert-g"}[typ]
            st.markdown(f'<div class="{cls}">{"🔴" if typ=="r" else "🟡" if typ=="y" else "✅"}<div><b>{titre}</b><br><small>{msg}</small></div></div>', unsafe_allow_html=True)

    # Export
    st.markdown("<br>", unsafe_allow_html=True)
    section("📥","Export Rapport Excel")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_fact.to_excel(w, sheet_name="Factures", index=False)
        df_inv.to_excel(w, sheet_name="Inventaire", index=False)
        df_ent.to_excel(w, sheet_name="Entrées Stock", index=False)
        df_sor.to_excel(w, sheet_name="Sorties Stock", index=False)
        df_prod.to_excel(w, sheet_name="Production", index=False)
        df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"].to_excel(w, sheet_name="Impayés", index=False)
    buf.seek(0)
    c_dl1, c_dl2 = st.columns([1, 3])
    with c_dl1:
        st.download_button(
            "⬇️ Télécharger Rapport Excel",
            data=buf,
            file_name=f"MULTIPACK_Rapport_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
