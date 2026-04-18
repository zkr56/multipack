"""
MULTIPACK SA – TABLEAU DE BORD DIRECTION GÉNÉRALE
Version Ultime v4 — Graphiques clairs, commentaires direction, prévisions & anticipations
Zone Industrielle de Yopougon, Abidjan, Côte d'Ivoire
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime, timedelta
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
import warnings
warnings.filterwarnings("ignore")
import io

st.set_page_config(page_title="MULTIPACK SA", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

C = {
    "bleu":"#6172F3","bleu_clair":"#C5CBFB","rose":"#FF6BA7","violet":"#51459E",
    "bleu_fonce":"#3C41CD","vert":"#2ECC71","rouge":"#E74C3C","orange":"#F39C12",
    "gris_bg":"#F4F6FB","blanc":"#FFFFFF","texte":"#2D3748","muted":"#718096",
    "bordure":"#E2E8F0","sidebar":"#1E2A4A","or":"#D69E2E",
}
PIE  = ["#6172F3","#FF6BA7","#51459E","#F39C12","#2ECC71","#3C41CD","#C5CBFB","#E74C3C"]
LINE = ["#6172F3","#FF6BA7","#51459E","#2ECC71","#F39C12"]

# ══════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
*,*::before,*::after{{box-sizing:border-box;}}
html,body,[class*="css"]{{font-family:'Inter',sans-serif;background:{C['gris_bg']};color:{C['texte']};}}
.main .block-container{{background:{C['gris_bg']};padding:1.2rem 2rem 2rem;max-width:1600px;}}
[data-testid="stSidebar"]{{background:{C['sidebar']} !important;border-right:none;}}
[data-testid="stSidebar"] *{{color:#E2E8F0 !important;}}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label{{
    background:rgba(255,255,255,0.04);border-radius:8px;padding:9px 14px !important;
    margin:3px 0 !important;border:1px solid transparent;font-size:0.85rem !important;
    font-weight:500 !important;color:#CBD5E1 !important;transition:all 0.15s;}}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover{{
    background:rgba(97,114,243,0.2) !important;border-color:rgba(97,114,243,0.5) !important;}}
.mp-header{{background:linear-gradient(135deg,#1E2A4A 0%,#2D3A6B 55%,{C['bleu']} 100%);
    border-radius:14px;padding:26px 34px;margin-bottom:22px;
    display:flex;justify-content:space-between;align-items:center;
    box-shadow:0 4px 20px rgba(97,114,243,0.22);}}
.mp-header h1{{font-size:1.65rem;font-weight:800;color:white;margin:0 0 4px;letter-spacing:-0.02em;}}
.mp-header .sub{{color:rgba(255,255,255,0.68);font-size:0.82rem;}}
.mp-header .badge{{background:rgba(255,255,255,0.13);border:1px solid rgba(255,255,255,0.22);
    border-radius:10px;padding:10px 20px;text-align:center;color:white;min-width:130px;}}
.mp-header .badge .day{{font-size:1.6rem;font-weight:800;display:block;line-height:1.1;}}
.kpi{{background:{C['blanc']};border-radius:12px;padding:18px 20px;border:1px solid {C['bordure']};
    box-shadow:0 1px 5px rgba(0,0,0,0.06);position:relative;overflow:hidden;transition:box-shadow 0.2s;height:100%;}}
.kpi:hover{{box-shadow:0 5px 18px rgba(0,0,0,0.10);}}
.kpi::before{{content:'';position:absolute;top:0;left:0;width:100%;height:4px;border-radius:12px 12px 0 0;}}
.kpi.bleu::before{{background:{C['bleu']};}} .kpi.vert::before{{background:{C['vert']};}}
.kpi.rouge::before{{background:{C['rouge']};}} .kpi.orange::before{{background:{C['orange']};}}
.kpi.violet::before{{background:{C['violet']};}} .kpi.rose::before{{background:{C['rose']};}}
.kpi.or::before{{background:{C['or']};}}
.kpi .ico{{font-size:1.4rem;margin-bottom:8px;display:block;}}
.kpi .lbl{{font-size:0.68rem;font-weight:700;color:{C['muted']};text-transform:uppercase;letter-spacing:0.07em;}}
.kpi .val{{font-size:1.55rem;font-weight:800;color:{C['texte']};margin:5px 0 3px;line-height:1.15;}}
.kpi .sub{{font-size:0.73rem;color:{C['muted']};}}
.sec{{font-size:0.8rem;font-weight:700;color:{C['muted']};text-transform:uppercase;
    letter-spacing:0.09em;padding-bottom:8px;margin:22px 0 10px;border-bottom:2px solid {C['bordure']};}}
.comment-box{{background:linear-gradient(135deg,#EEF2FF,#F0F4FF);border-left:4px solid {C['bleu']};
    border-radius:0 10px 10px 0;padding:13px 16px;margin:8px 0 14px;font-size:0.84rem;color:#374151;line-height:1.55;}}
.comment-box .ct{{font-weight:700;color:{C['bleu']};font-size:0.78rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:5px;}}
.prevision-box{{background:linear-gradient(135deg,#FFFBEB,#FEF3C7);border-left:4px solid {C['orange']};
    border-radius:0 10px 10px 0;padding:13px 16px;margin:8px 0 14px;font-size:0.84rem;color:#374151;line-height:1.55;}}
.prevision-box .ct{{font-weight:700;color:{C['orange']};font-size:0.78rem;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:5px;}}
.alert-r{{background:#FFF5F5;border-left:4px solid {C['rouge']};border-radius:0 8px 8px 0;
    padding:11px 15px;margin:5px 0;font-size:0.83rem;color:#7B1C1C;display:flex;gap:9px;align-items:flex-start;}}
.alert-y{{background:#FFFBEB;border-left:4px solid {C['orange']};border-radius:0 8px 8px 0;
    padding:11px 15px;margin:5px 0;font-size:0.83rem;color:#78350F;display:flex;gap:9px;align-items:flex-start;}}
.alert-g{{background:#F0FFF4;border-left:4px solid {C['vert']};border-radius:0 8px 8px 0;
    padding:11px 15px;margin:5px 0;font-size:0.83rem;color:#14532D;display:flex;gap:9px;align-items:flex-start;}}
.prev-kpi{{background:linear-gradient(135deg,#1E2A4A,#2D3A6B);border-radius:12px;padding:18px 20px;
    border:1px solid rgba(97,114,243,0.3);color:white;height:100%;}}
.prev-kpi .ico{{font-size:1.4rem;margin-bottom:8px;display:block;}}
.prev-kpi .lbl{{font-size:0.68rem;font-weight:700;color:rgba(255,255,255,0.6);text-transform:uppercase;letter-spacing:0.07em;}}
.prev-kpi .val{{font-size:1.5rem;font-weight:800;color:white;margin:5px 0 3px;}}
.prev-kpi .sub{{font-size:0.73rem;color:rgba(255,255,255,0.55);}}
.prog{{margin:7px 0;}} .prog .pl{{display:flex;justify-content:space-between;font-size:0.74rem;color:{C['muted']};margin-bottom:4px;}}
.prog .pb{{height:8px;background:{C['bordure']};border-radius:4px;overflow:hidden;}}
.prog .pf{{height:100%;border-radius:4px;transition:width 0.4s;}}
#MainMenu,footer{{visibility:hidden;}}.stDeployButton{{display:none;}}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════
def excel_style(fig, height=320, show_legend=True):
    fig.update_layout(
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, sans-serif", color=C["texte"], size=11),
        height=height, margin=dict(t=20, b=40, l=10, r=20),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=C["texte"], size=10),
                    orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=dict(bgcolor="white", font_size=12, bordercolor=C["bordure"]),
        showlegend=show_legend,
    )
    fig.update_xaxes(gridcolor="#EAECF0", showgrid=True, zeroline=False,
                     linecolor=C["bordure"], tickfont=dict(color=C["muted"], size=10))
    fig.update_yaxes(gridcolor="#EAECF0", showgrid=True, zeroline=False,
                     linecolor=C["bordure"], tickfont=dict(color=C["muted"], size=10))
    return fig

def kpi(icon, label, value, sub="", color="bleu"):
    st.markdown(f"""<div class="kpi {color}">
        <span class="ico">{icon}</span><div class="lbl">{label}</div>
        <div class="val">{value}</div><div class="sub">{sub}</div></div>""", unsafe_allow_html=True)

def prev_kpi(icon, label, value, sub=""):
    st.markdown(f"""<div class="prev-kpi">
        <span class="ico">{icon}</span><div class="lbl">{label}</div>
        <div class="val">{value}</div><div class="sub">{sub}</div></div>""", unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="sec">{title}</div>', unsafe_allow_html=True)

def comment(texte, titre="📌 Ce que ça signifie pour vous"):
    st.markdown(f"""<div class="comment-box"><div class="ct">{titre}</div>{texte}</div>""", unsafe_allow_html=True)

def prevision_comment(texte, titre="🔮 Prévision & Anticipation"):
    st.markdown(f"""<div class="prevision-box"><div class="ct">{titre}</div>{texte}</div>""", unsafe_allow_html=True)

def fmt(v, suffix="FCFA"):
    if v >= 1_000_000_000: return f"{v/1_000_000_000:.2f} Md {suffix}"
    if v >= 1_000_000:     return f"{v/1_000_000:.2f} M {suffix}"
    if v >= 1_000:         return f"{v/1_000:.1f} K {suffix}"
    return f"{v:,.0f} {suffix}"

def prog_bar(label, value, max_val, color=None):
    color = color or C["bleu"]
    pct = min(100, value / max_val * 100) if max_val else 0
    st.markdown(f"""<div class="prog">
        <div class="pl"><span>{label}</span><span><b>{value:,.0f} FCFA</b></span></div>
        <div class="pb"><div class="pf" style="width:{pct:.1f}%;background:{color};"></div></div>
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# MODÈLE DE PRÉVISION
# ══════════════════════════════════════════════
def prevoir_ca(serie_mensuelle: pd.Series, n_mois: int = 6):
    """
    Régression polynomiale deg=2 sur série temporelle.
    Retourne : df avec colonnes Période, CA_Prévu, Bas, Haut
    """
    y = serie_mensuelle.values.astype(float)
    X = np.arange(len(y)).reshape(-1, 1)
    poly = PolynomialFeatures(degree=2)
    Xp = poly.fit_transform(X)
    model = LinearRegression().fit(Xp, y)
    residuals = y - model.predict(Xp)
    std = residuals.std()

    X_fut = np.arange(len(y), len(y) + n_mois).reshape(-1, 1)
    y_fut = model.predict(poly.transform(X_fut))
    y_fut = np.maximum(y_fut, 0)

    last_date = serie_mensuelle.index[-1]
    periodes = [str(pd.Period(last_date, "M") + i + 1) for i in range(n_mois)]
    return pd.DataFrame({
        "Période": periodes,
        "CA_Prévu": y_fut,
        "Bas": np.maximum(y_fut - 1.5 * std, 0),
        "Haut": y_fut + 1.5 * std,
    })


def prevoir_stock(stock_final, sorties_mensuelles, seuil, n_mois=6):
    """
    Prédit le nombre de mois avant rupture de stock.
    """
    if sorties_mensuelles <= 0:
        return n_mois, [stock_final] * n_mois
    projection = [max(0, stock_final - sorties_mensuelles * (i+1)) for i in range(n_mois)]
    mois_rupture = next((i+1 for i, v in enumerate(projection) if v <= seuil), None)
    return mois_rupture, projection


# ══════════════════════════════════════════════
# GÉNÉRATION DES DONNÉES
# ══════════════════════════════════════════════


# ══════════════════════════════════════════════
# GÉNÉRATION DES DONNÉES — MULTIPACK v5
# 13 segments · 60+ clients · 8 zones · 1200 factures
# ══════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def generer_donnees():
    import random as rnd
    rnd.seed(42)
    np.random.seed(42)

    PRODUITS = [
        # Sacs supermarché
        {"ref":"SAC-SM-001","designation":"Sac Supermarché 30L Standard",      "categorie":"Sacs Supermarché",        "prix_achat":1800,  "prix_vente":2500,  "seuil":200},
        {"ref":"SAC-SM-002","designation":"Sac Supermarché 50L Renforcé",       "categorie":"Sacs Supermarché",        "prix_achat":2600,  "prix_vente":3500,  "seuil":200},
        {"ref":"SAC-SM-003","designation":"Sac Supermarché 20L Mini",           "categorie":"Sacs Supermarché",        "prix_achat":1400,  "prix_vente":2000,  "seuil":300},
        {"ref":"SAC-SM-004","designation":"Sac Supermarché HD Noir 40L",        "categorie":"Sacs Supermarché",        "prix_achat":2200,  "prix_vente":3000,  "seuil":150},
        {"ref":"SAC-SM-005","designation":"Sac Caisse Transparent 25L",         "categorie":"Sacs Supermarché",        "prix_achat":4500,  "prix_vente":6200,  "seuil":100},
        {"ref":"SAC-SM-006","designation":"Sac Biosourcé Compostable 30L",      "categorie":"Sacs Supermarché",        "prix_achat":5200,  "prix_vente":7500,  "seuil":80},
        # Sacs poubelles
        {"ref":"SAC-PB-001","designation":"Sac Poubelle 30L Vert",              "categorie":"Sacs Poubelles",          "prix_achat":650,   "prix_vente":950,   "seuil":500},
        {"ref":"SAC-PB-002","designation":"Sac Poubelle 50L Noir Renforcé",     "categorie":"Sacs Poubelles",          "prix_achat":850,   "prix_vente":1200,  "seuil":500},
        {"ref":"SAC-PB-003","designation":"Sac Poubelle 110L Industrial",       "categorie":"Sacs Poubelles",          "prix_achat":1200,  "prix_vente":1700,  "seuil":300},
        {"ref":"SAC-PB-004","designation":"Sac Poubelle 20L Bleu Ménager",      "categorie":"Sacs Poubelles",          "prix_achat":550,   "prix_vente":800,   "seuil":400},
        {"ref":"SAC-PB-005","designation":"Sac Poubelle 240L Collectivité",     "categorie":"Sacs Poubelles",          "prix_achat":2200,  "prix_vente":3000,  "seuil":100},
        # Gobelets
        {"ref":"GOB-001",   "designation":"Gobelet Plastique 20cl Standard",    "categorie":"Gobelets",                "prix_achat":3500,  "prix_vente":5000,  "seuil":150},
        {"ref":"GOB-002",   "designation":"Gobelet Plastique 33cl Grand",       "categorie":"Gobelets",                "prix_achat":2800,  "prix_vente":4000,  "seuil":150},
        {"ref":"GOB-003",   "designation":"Gobelet Cristal 25cl",               "categorie":"Gobelets",                "prix_achat":3200,  "prix_vente":4500,  "seuil":100},
        {"ref":"GOB-004",   "designation":"Gobelet Solo Cup 50cl",              "categorie":"Gobelets",                "prix_achat":2400,  "prix_vente":3500,  "seuil":80},
        {"ref":"GOB-005",   "designation":"Gobelet Dégustation 10cl",           "categorie":"Gobelets",                "prix_achat":3800,  "prix_vente":5500,  "seuil":120},
        # Emballages alimentaires
        {"ref":"EMB-AL-001","designation":"Barquette Alimentaire PP 500ml",     "categorie":"Emballages Alimentaires", "prix_achat":4200,  "prix_vente":6000,  "seuil":100},
        {"ref":"EMB-AL-002","designation":"Barquette Alimentaire PP 1L",        "categorie":"Emballages Alimentaires", "prix_achat":3600,  "prix_vente":5200,  "seuil":80},
        {"ref":"EMB-AL-003","designation":"Film Étirable Alimentaire 300m",     "categorie":"Emballages Alimentaires", "prix_achat":5500,  "prix_vente":8000,  "seuil":50},
        {"ref":"EMB-AL-004","designation":"Sachet Zip PP 1L",                   "categorie":"Emballages Alimentaires", "prix_achat":1800,  "prix_vente":2600,  "seuil":200},
        # Emballages industriels
        {"ref":"EMB-IN-001","designation":"Big Bag Industriel 1T FIBC",         "categorie":"Emballages Industriels",  "prix_achat":8500,  "prix_vente":12000, "seuil":30},
        {"ref":"EMB-IN-002","designation":"Fût Plastique HDPE 220L",            "categorie":"Emballages Industriels",  "prix_achat":12000, "prix_vente":17000, "seuil":20},
        {"ref":"EMB-IN-003","designation":"Jerrycan HDPE 10L",                  "categorie":"Emballages Industriels",  "prix_achat":2800,  "prix_vente":4000,  "seuil":60},
        {"ref":"EMB-IN-004","designation":"Bidon HDPE 5L avec bouchon",         "categorie":"Emballages Industriels",  "prix_achat":1600,  "prix_vente":2400,  "seuil":80},
        {"ref":"EMB-IN-005","designation":"Sac Industriel PE 50Kg",             "categorie":"Emballages Industriels",  "prix_achat":6500,  "prix_vente":9000,  "seuil":40},
        # Matières premières
        {"ref":"MAT-PE-001","designation":"Granulés PE-HD",                     "categorie":"Matières Premières",      "prix_achat":14000, "prix_vente":None,  "seuil":100},
        {"ref":"MAT-PP-001","designation":"Granulés PP",                        "categorie":"Matières Premières",      "prix_achat":15500, "prix_vente":None,  "seuil":80},
        {"ref":"MAT-PVC-001","designation":"Granulés PVC Souple",               "categorie":"Matières Premières",      "prix_achat":18000, "prix_vente":None,  "seuil":60},
        {"ref":"MAT-REC-001","designation":"Plastique Recyclé Mix",             "categorie":"Matières Premières",      "prix_achat":7500,  "prix_vente":None,  "seuil":50},
        {"ref":"MAT-COL-001","designation":"Masterbatch Noir",                  "categorie":"Matières Premières",      "prix_achat":22000, "prix_vente":None,  "seuil":40},
        {"ref":"MAT-COL-002","designation":"Masterbatch Vert",                  "categorie":"Matières Premières",      "prix_achat":23000, "prix_vente":None,  "seuil":35},
    ]

    # ─── 13 SEGMENTS · 60+ CLIENTS · 8 ZONES ──────────────────────
    CLIENTS = [
        # GRANDE DISTRIBUTION (zone Sud/Nord)
        {"nom":"SOCOCÉ Abidjan Plateau",         "segment":"Grande Distribution",      "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Carrefour Market Cocody",         "segment":"Grande Distribution",      "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"PROSUMA SA",                      "segment":"Grande Distribution",      "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Leader Price Yopougon",           "segment":"Grande Distribution",      "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"CDCI Supermarché Plateau",        "segment":"Grande Distribution",      "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"Hypermarché Prima",               "segment":"Grande Distribution",      "zone":"Zone Sud",    "ville":"Abidjan"},
        # INDUSTRIEL AGROALIMENTAIRE
        {"nom":"NESTLE CI",                       "segment":"Industriel Agroalimentaire","zone":"Zone Sud",   "ville":"Abidjan"},
        {"nom":"SIC CACAOS",                      "segment":"Industriel Agroalimentaire","zone":"Zone Sud",   "ville":"Abidjan"},
        {"nom":"SIFCA Groupe",                    "segment":"Industriel Agroalimentaire","zone":"Zone Sud",   "ville":"Abidjan"},
        {"nom":"PALMCI San-Pédro",                "segment":"Industriel Agroalimentaire","zone":"Zone Ouest", "ville":"San-Pédro"},
        {"nom":"Ivoire Sucre",                    "segment":"Industriel Agroalimentaire","zone":"Zone Centre","ville":"Bouaké"},
        {"nom":"BLOHORN SA",                      "segment":"Industriel Agroalimentaire","zone":"Zone Sud",   "ville":"Abidjan"},
        # INDUSTRIEL MANUFACTURIER
        {"nom":"SOLIBRA Brasserie",               "segment":"Industriel Manufacturier", "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"GONFREVILLE Textile",             "segment":"Industriel Manufacturier", "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"CIE Côte d'Ivoire",               "segment":"Industriel Manufacturier", "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"SITARAIL CI",                     "segment":"Industriel Manufacturier", "zone":"Zone Centre", "ville":"Bouaké"},
        # DISTRIBUTEUR LOCAL ABIDJAN
        {"nom":"Commerce Général Yop",            "segment":"Distributeur Local Abidjan","zone":"Zone Nord",  "ville":"Abidjan"},
        {"nom":"Maquis & Resto Supplies",         "segment":"Distributeur Local Abidjan","zone":"Zone Sud",   "ville":"Abidjan"},
        {"nom":"NEGOCE MARCORY",                  "segment":"Distributeur Local Abidjan","zone":"Zone Sud",   "ville":"Abidjan"},
        {"nom":"DIST ABOBO EXPRESS",              "segment":"Distributeur Local Abidjan","zone":"Zone Nord",  "ville":"Abidjan"},
        {"nom":"COMMERCE ADJAMÉ",                 "segment":"Distributeur Local Abidjan","zone":"Zone Nord",  "ville":"Abidjan"},
        # DISTRIBUTEUR RÉGIONAL CI
        {"nom":"DIOULA NÉGOCE Bouaké",            "segment":"Distributeur Régional CI", "zone":"Zone Centre", "ville":"Bouaké"},
        {"nom":"AGRO DIST Korhogo",               "segment":"Distributeur Régional CI", "zone":"Zone Nord",   "ville":"Korhogo"},
        {"nom":"TRANSIT SAN-PEDRO",               "segment":"Distributeur Régional CI", "zone":"Zone Ouest",  "ville":"San-Pédro"},
        {"nom":"COMMERCE ABENGOUROU",             "segment":"Distributeur Régional CI", "zone":"Zone Est",    "ville":"Abengourou"},
        {"nom":"DIST DALOA CENTER",               "segment":"Distributeur Régional CI", "zone":"Zone Centre", "ville":"Daloa"},
        {"nom":"NÉGOCE MAN OUEST",                "segment":"Distributeur Régional CI", "zone":"Zone Ouest",  "ville":"Man"},
        # HÔTELLERIE & RESTAURATION
        {"nom":"SOFITEL Abidjan",                 "segment":"Hôtellerie & Restauration","zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Hôtel Ivoire Intercontinental",   "segment":"Hôtellerie & Restauration","zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Groupe LAICO",                    "segment":"Hôtellerie & Restauration","zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Palm Club Hôtel",                 "segment":"Hôtellerie & Restauration","zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Restaurant Le Phare",             "segment":"Hôtellerie & Restauration","zone":"Zone Sud",    "ville":"Abidjan"},
        # SANTÉ & PHARMACIE
        {"nom":"CHU de Treichville",              "segment":"Santé & Pharmacie",        "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Clinique Biasa",                  "segment":"Santé & Pharmacie",        "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"LABOREX CI",                      "segment":"Santé & Pharmacie",        "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Pharmivoire Groupe",              "segment":"Santé & Pharmacie",        "zone":"Zone Sud",    "ville":"Abidjan"},
        # ADMINISTRATION & COLLECTIVITÉS
        {"nom":"Mairie Cocody",                   "segment":"Administration & Collectivités","zone":"Zone Sud","ville":"Abidjan"},
        {"nom":"Mairie de Bouaké",                "segment":"Administration & Collectivités","zone":"Zone Centre","ville":"Bouaké"},
        {"nom":"ANADER CI",                       "segment":"Administration & Collectivités","zone":"Zone Sud","ville":"Abidjan"},
        {"nom":"ONAD (Assainissement)",           "segment":"Administration & Collectivités","zone":"Zone Sud","ville":"Abidjan"},
        # BTP & CONSTRUCTION
        {"nom":"SOGEBAC Construction",            "segment":"BTP & Construction",       "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"COLAS CI",                        "segment":"BTP & Construction",       "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"BATIPRO Group",                   "segment":"BTP & Construction",       "zone":"Zone Nord",   "ville":"Abidjan"},
        # LOGISTIQUE & TRANSPORT
        {"nom":"BOLLORÉ Logistics CI",            "segment":"Logistique & Transport",   "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"SDV Port Abidjan",                "segment":"Logistique & Transport",   "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"MAERSK CI",                       "segment":"Logistique & Transport",   "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"TransCi Freight",                 "segment":"Logistique & Transport",   "zone":"Zone Ouest",  "ville":"San-Pédro"},
        # AGRICULTURE & AGRO-INDUSTRIE
        {"nom":"SOGB Rubber",                     "segment":"Agriculture & Agro-Industrie","zone":"Zone Ouest","ville":"San-Pédro"},
        {"nom":"SAPH Hévéa CI",                   "segment":"Agriculture & Agro-Industrie","zone":"Zone Ouest","ville":"Aboisso"},
        {"nom":"COOPAGRI Daloa",                  "segment":"Agriculture & Agro-Industrie","zone":"Zone Centre","ville":"Daloa"},
        # EXPORT SOUS-RÉGIONAL
        {"nom":"WEST AFRICA TRADING CO.",         "segment":"Export Sous-Régional",     "zone":"Export",      "ville":"Dakar"},
        {"nom":"GHANA IMPORT GROUP",              "segment":"Export Sous-Régional",     "zone":"Export",      "ville":"Accra"},
        {"nom":"MALI PACKAGING SARL",             "segment":"Export Sous-Régional",     "zone":"Export",      "ville":"Bamako"},
        {"nom":"BURKINA EMBALLAGE",               "segment":"Export Sous-Régional",     "zone":"Export",      "ville":"Ouagadougou"},
        {"nom":"TOGO DISTRIBUTION",               "segment":"Export Sous-Régional",     "zone":"Export",      "ville":"Lomé"},
        # ARTISANAT & PME
        {"nom":"COOPEC Artisans Yop",             "segment":"Artisanat & PME",          "zone":"Zone Nord",   "ville":"Abidjan"},
        {"nom":"PME Couture Mode CI",             "segment":"Artisanat & PME",          "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"Traiteur Événements Prestige",    "segment":"Artisanat & PME",          "zone":"Zone Sud",    "ville":"Abidjan"},
        {"nom":"MARCHÉ ADJAMÉ GROSSISTE",         "segment":"Artisanat & PME",          "zone":"Zone Nord",   "ville":"Abidjan"},
    ]

    MODES = ["Virement bancaire","Chèque certifié","Espèce","Carte bancaire","Traite"]
    POIDS = [0.40, 0.25, 0.18, 0.10, 0.07]
    SAISON= [0.75,0.70,0.80,0.85,0.90,0.97,0.88,1.12,1.22,1.30,1.45,1.38]

    SEG_W = {
        "Grande Distribution":3.5, "Industriel Agroalimentaire":3.2,
        "Industriel Manufacturier":2.8, "Distributeur Local Abidjan":1.8,
        "Distributeur Régional CI":1.5, "Hôtellerie & Restauration":1.2,
        "Santé & Pharmacie":1.0, "Administration & Collectivités":1.3,
        "BTP & Construction":1.6, "Logistique & Transport":1.4,
        "Agriculture & Agro-Industrie":2.0, "Export Sous-Régional":2.5,
        "Artisanat & PME":0.8,
    }
    P_PAY = {
        "Grande Distribution":0.90, "Industriel Agroalimentaire":0.87,
        "Industriel Manufacturier":0.85, "Distributeur Local Abidjan":0.75,
        "Distributeur Régional CI":0.70, "Hôtellerie & Restauration":0.92,
        "Santé & Pharmacie":0.95, "Administration & Collectivités":0.60,
        "BTP & Construction":0.72, "Logistique & Transport":0.88,
        "Agriculture & Agro-Industrie":0.80, "Export Sous-Régional":0.78,
        "Artisanat & PME":0.65,
    }

    start = datetime(2022, 1, 1)
    pv = [p for p in PRODUITS if p["prix_vente"]]
    factures = []
    for i in range(1200):  # 1200 factures vs 850 avant
        cl  = rnd.choice(CLIENTS)
        d   = start + timedelta(days=rnd.randint(0, 1094))
        co  = SAISON[d.month-1] * SEG_W.get(cl["segment"], 1.5)
        nb  = rnd.choices([1,2,3,4], weights=[0.25,0.38,0.27,0.10])[0]
        sel = rnd.sample(pv, min(nb, len(pv)))
        mht = round(sum(p["prix_vente"] * rnd.randint(1, int(10*co)) for p in sel), 0)
        tva = round(mht * 0.18, 0)
        ttc = mht + tva
        ok  = rnd.random() < P_PAY.get(cl["segment"], 0.80)
        mode = rnd.choices(MODES, weights=POIDS)[0] if ok else ""
        factures.append({
            "N° FACTURE":       f"F{d.year}-{str(d.month).zfill(2)}-{str(i+1).zfill(4)}",
            "DATE":             d, "ANNEE": d.year, "MOIS": d.month,
            "CLIENT":           cl["nom"], "SEGMENT": cl["segment"],
            "ZONE":             cl["zone"], "VILLE": cl["ville"],
            "MONTANT HT":       mht, "TVA": tva, "MONTANT TTC": ttc,
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
            d = start + timedelta(days=rnd.randint(0, 1094))
            qte = rnd.randint(10, 600)
            cout= round(p["prix_achat"] * rnd.uniform(0.91, 1.06), 0)
            ent_rows.append({"Date":d,"Référence":p["ref"],"Désignation":p["designation"],
                             "Catégorie":p["categorie"],"Coût d'achat":cout,"Quantité":qte,"Total":round(cout*qte,0)})
        if p["prix_vente"]:
            for _ in range(rnd.randint(8, 40)):
                d = start + timedelta(days=rnd.randint(0, 1094))
                qte = rnd.randint(5, 350)
                pv_ = round(p["prix_vente"] * rnd.uniform(0.97, 1.10), 0)
                sor_rows.append({"Date":d,"Référence":p["ref"],"Désignation":p["designation"],
                                 "Catégorie":p["categorie"],"Prix de vente":pv_,"Quantité":qte,"Total":round(pv_*qte,0)})
    df_ent = pd.DataFrame(ent_rows).sort_values("Date").reset_index(drop=True)
    df_sor = pd.DataFrame(sor_rows).sort_values("Date").reset_index(drop=True)

    for i, row in df_inv.iterrows():
        ref = row["ref"]
        e = df_ent[df_ent["Référence"]==ref]["Quantité"].sum()
        s = df_sor[df_sor["Référence"]==ref]["Quantité"].sum() if ref in df_sor["Référence"].values else 0
        sf= max(0, int(row["Stock initial"]) + int(e) - int(s))
        df_inv.at[i,"Entrées"] = int(e); df_inv.at[i,"Sorties"] = int(s)
        df_inv.at[i,"Stock final"] = sf
        df_inv.at[i,"Valeur"] = round(sf * row["prix_achat"], 0)
        df_inv.at[i,"Sorties_moy_mois"] = round(s / 36, 1)
        df_inv.at[i,"Statut"] = ("Non disponible" if sf==0 else "Stock faible" if sf<row["seuil"] else "Stock normal")

    machines = ["Presse 1 – Injection","Presse 2 – Soufflage","Extrudeuse A","Extrudeuse B","Thermoformeuse"]
    prod_rows = []
    for m in range(36):
        d = start + pd.DateOffset(months=m)
        co = SAISON[d.month-1]
        for mach in machines:
            r  = rnd.uniform(0.73, 0.97); pp = int(rnd.randint(50000,140000)*co)
            pr = int(pp*r); reb = int(pr*rnd.uniform(0.01,0.055))
            prod_rows.append({
                "Mois":d.strftime("%Y-%m"),"Mois_Label":d.strftime("%b %Y"),"Année":d.year,"Machine":mach,
                "Production planifiée":pp,"Production réelle":pr,"Rebuts":reb,
                "Taux rendement":round(r*100,1),"Taux rebut":round(reb/pr*100,2) if pr>0 else 0,
            })
    df_prod = pd.DataFrame(prod_rows)
    return df_f, df_inv, df_ent, df_sor, df_prod


# ══════════════════════════════════════════════
# CHARGEMENT & SIDEBAR v5
# ══════════════════════════════════════════════
with st.spinner("Chargement du tableau de bord…"):
    df_fact, df_inv, df_ent, df_sor, df_prod = generer_donnees()

ALL_SEGMENTS = sorted(df_fact["SEGMENT"].unique().tolist())
ALL_ZONES    = sorted(df_fact["ZONE"].unique().tolist())
ALL_ANNEES   = sorted(df_fact["ANNEE"].unique().tolist())

with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:20px 0 12px;">
        <div style="font-size:2.2rem;">📦</div>
        <div style="font-size:1.05rem;font-weight:800;color:white;">MULTIPACK SA</div>
        <div style="font-size:0.68rem;color:#94A3B8;margin-top:2px;">Zone Ind. Yopougon · Abidjan</div>
    </div>
    <hr style="border-color:rgba(255,255,255,0.1);margin:6px 0 14px;">
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", [
        "🏠  Vue d'ensemble",
        "💰  Chiffre d'Affaires",
        "🛍️  Nos Produits",
        "👥  Nos Clients",
        "📦  Stocks & Inventaire",
        "🏭  Production",
        "💳  Paiements",
        "🔮  Prévisions & Anticipations",
        "📊  Comparaisons & Analyses",
        "⚠️  Alertes & Conseils",
    ], label_visibility="collapsed")

    st.markdown("<hr style='border-color:rgba(255,255,255,0.1);margin:10px 0;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.67rem;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:7px;padding-left:3px;'>🗓 Période analysée</div>", unsafe_allow_html=True)
    sel_annees = st.multiselect("Années", ALL_ANNEES, default=ALL_ANNEES)

    st.markdown("<div style='font-size:0.67rem;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin:10px 0 7px;padding-left:3px;'>🏷 Segments actifs</div>", unsafe_allow_html=True)
    sel_seg = st.multiselect("Segments", ALL_SEGMENTS, default=ALL_SEGMENTS,
                              help="Sélectionnez les segments à inclure dans toutes les pages")

    st.markdown("<div style='font-size:0.67rem;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.08em;margin:10px 0 7px;padding-left:3px;'>🗺 Zones actives</div>", unsafe_allow_html=True)
    sel_zones = st.multiselect("Zones", ALL_ZONES, default=ALL_ZONES,
                                help="Filtrez par zone géographique")

    st.markdown(f"<div style='margin-top:16px;font-size:0.67rem;color:#475569;text-align:center;line-height:1.7;'>Mise à jour : {datetime.now().strftime('%d/%m/%Y %H:%M')}<br>{len(df_fact)} factures · {len(ALL_SEGMENTS)} segments · {len(df_fact['CLIENT'].unique())} clients<br>© MULTIPACK SA 2024</div>", unsafe_allow_html=True)

# Filtre global
df_f = df_fact[
    df_fact["ANNEE"].isin(sel_annees) &
    df_fact["SEGMENT"].isin(sel_seg) &
    df_fact["ZONE"].isin(sel_zones)
]

now = datetime.now()
st.markdown(f"""
<div class="mp-header">
  <div>
    <h1>📦 MULTIPACK SA — Tableau de Bord</h1>
    <p class="sub">Pilotage Commercial &amp; Reporting · {len(ALL_SEGMENTS)} segments · {len(df_fact['CLIENT'].unique())} clients · {len(df_fact)} factures · Direction Générale</p>
  </div>
  <div class="badge"><span class="day">{now.strftime('%d')}</span>{now.strftime('%b %Y')}<br>
  <span style="opacity:0.65;font-size:0.68rem;">{now.strftime('%H:%M')}</span></div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# PAGE 1 : VUE D'ENSEMBLE
# ══════════════════════════════════════════════════════════════
if page == "🏠  Vue d'ensemble":
    ca_total = df_f["MONTANT TTC"].sum()
    ca_paye  = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    ca_impaye= df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]["MONTANT TTC"].sum()
    taux_rec = ca_paye/ca_total*100 if ca_total else 0
    val_stock= df_inv["Valeur"].sum()
    nb_rupt  = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib  = (df_inv["Statut"]=="Stock faible").sum()
    rend_moy = df_prod["Taux rendement"].mean()

    comment(f"Bonjour. Voici la situation globale de MULTIPACK SA. Sur la période sélectionnée, "
            f"nous avons réalisé <b>{fmt(ca_total)}</b> de chiffre d'affaires, dont <b>{fmt(ca_paye)}</b> encaissés. "
            f"Il reste <b>{fmt(ca_impaye)}</b> à recouvrer. Nos stocks valent <b>{fmt(val_stock)}</b> "
            f"et nos machines tournent à <b>{rend_moy:.0f}%</b> de rendement.", "📋 Résumé exécutif")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("💰","CA Total", fmt(ca_total), f"{len(df_f)} factures","bleu")
    with c2: kpi("✅","Encaissé", fmt(ca_paye), f"Taux : {taux_rec:.1f}%","vert")
    with c3: kpi("⏳","Impayés", fmt(ca_impaye), f"{len(df_f[df_f['ETAT DE PAIEMENT']=='Impayée'])} factures","rouge")
    with c4: kpi("👥","Clients Actifs", str(df_f["CLIENT"].nunique()), f"{df_f['SEGMENT'].nunique()} segments","violet")
    st.markdown("<br>", unsafe_allow_html=True)
    c5,c6,c7,c8 = st.columns(4)
    with c5: kpi("🏭","Rendement Usine", f"{rend_moy:.1f}%","Moyenne machines","bleu")
    with c6: kpi("📦","Valeur Stock", fmt(val_stock), f"{len(df_inv)} références","violet")
    with c7: kpi("🟡","Stocks Faibles", str(nb_faib),"À commander bientôt","orange")
    with c8: kpi("🔴","Ruptures", str(nb_rupt),"Produits épuisés — urgent","rouge" if nb_rupt>0 else "vert")
    st.markdown("<br>", unsafe_allow_html=True)

    # Courbe CA mensuelle seule
    section("📈  ÉVOLUTION DU CHIFFRE D'AFFAIRES — Mois par mois")
    comment("Chaque point de cette courbe = les ventes d'un mois. Une courbe qui monte = bonne dynamique. "
            "La ligne pointillée = notre objectif mensuel moyen à dépasser.")
    ca_m = df_f.copy(); ca_m["Période"] = ca_m["DATE"].dt.to_period("M").astype(str)
    ca_mois = ca_m.groupby("Période")["MONTANT TTC"].sum().reset_index(); ca_mois.columns = ["Période","CA"]
    objectif = ca_mois["CA"].mean() * 1.10
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=ca_mois["Période"], y=ca_mois["CA"], mode="lines+markers",
        name="CA mensuel", line=dict(color=C["bleu"], width=3),
        marker=dict(size=8, color=C["bleu"], line=dict(color="white",width=2)),
        fill="tozeroy", fillcolor="rgba(97,114,243,0.10)",
        hovertemplate="<b>%{x}</b><br>CA : %{y:,.0f} FCFA<extra></extra>"))
    fig.add_hline(y=objectif, line_dash="dot", line_color=C["rose"], line_width=2,
        annotation_text=f"Objectif {fmt(objectif,'').strip()}", annotation_font_color=C["rose"],
        annotation_position="top right")
    fig = excel_style(fig, 320, False)
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=9))
    fig.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("🥧  PART DES PAIEMENTS REÇUS")
        comment(f"Sur <b>{fmt(ca_total)}</b> de ventes, nous avons encaissé <b>{taux_rec:.0f}%</b>. "
                f"Un bon taux se situe au-dessus de 80%.")
        stat_d = df_f.groupby("ETAT DE PAIEMENT")["MONTANT TTC"].sum().reset_index()
        clrs = {r: (C["bleu"] if r=="Payée" else C["rouge"]) for r in stat_d["ETAT DE PAIEMENT"]}
        fig2 = go.Figure(go.Pie(
            labels=stat_d["ETAT DE PAIEMENT"], values=stat_d["MONTANT TTC"], hole=0.65,
            marker=dict(colors=[clrs[r] for r in stat_d["ETAT DE PAIEMENT"]], line=dict(color="white",width=3)),
            textinfo="label+percent", textfont=dict(size=12),
            hovertemplate="%{label}<br><b>%{value:,.0f} FCFA</b> (%{percent})<extra></extra>"))
        fig2.update_layout(paper_bgcolor="#FFFFFF", height=300,
            margin=dict(t=10,b=30,l=10,r=10),
            legend=dict(orientation="h", y=-0.08, xanchor="center", x=0.5, font=dict(color=C["texte"],size=11)),
            annotations=[dict(text=f"<b>{taux_rec:.0f}%</b>", x=0.5, y=0.5, showarrow=False,
                              font=dict(size=22, color=C["bleu"]), xanchor="center")])
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🏆  TOP 5 CLIENTS")
        comment("Nos 5 meilleurs clients par chiffre d'affaires. Ce sont eux qu'il faut fidéliser en priorité.")
        top5 = df_f.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=False).head(5).reset_index()
        fig3 = go.Figure(go.Bar(
            x=top5["CLIENT"], y=top5["MONTANT TTC"],
            marker_color=[C["bleu"],C["bleu_clair"],C["violet"],C["rose"],C["bleu_fonce"]],
            text=top5["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
            textposition="outside", textfont=dict(size=11),
            hovertemplate="<b>%{x}</b><br>CA : %{y:,.0f} FCFA<extra></extra>"))
        fig3 = excel_style(fig3, 300, False)
        fig3.update_xaxes(tickangle=-20, tickfont=dict(size=10))
        fig3.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE 2 : CHIFFRE D'AFFAIRES
# ══════════════════════════════════════════════════════════════
elif page == "💰  Chiffre d'Affaires":
    ca = df_f["MONTANT TTC"].sum()
    paye = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    taux = paye/ca*100 if ca else 0
    ca21 = df_fact[df_fact["ANNEE"]==2022]["MONTANT TTC"].sum()
    ca22 = df_fact[df_fact["ANNEE"]==2023]["MONTANT TTC"].sum()
    crois = (ca22-ca21)/ca21*100 if ca21 else 0

    comment(f"CA total : <b>{fmt(ca)}</b>. Croissance 2022→2023 : <b>{crois:+.1f}%</b>. "
            f"Analysez les mois forts et faibles pour mieux planifier.", "📊 Guide de lecture")

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","CA Total", fmt(ca), f"{len(df_f)} factures","bleu")
    with c2: kpi("✅","Encaissé", fmt(paye), f"{taux:.1f}%","vert")
    with c3: kpi("🛒","Commande Moy.", fmt(df_f["MONTANT TTC"].mean()),"Par facture","violet")
    with c4: kpi("👥","Nb Clients", str(df_f["CLIENT"].nunique()),"Actifs","orange")
    with c5: kpi("📈","Croissance 22→23", f"{crois:+.1f}%","Variation annuelle","vert" if crois>0 else "rouge")
    st.markdown("<br>", unsafe_allow_html=True)

    # BARRES mensuelles colorées
    section("📊  CA PAR MOIS — Chaque barre = un mois de ventes")
    comment("🟢 Vert = mois fort &nbsp; 🟡 Orange = mois moyen &nbsp; 🔴 Rouge = mois faible. "
            "Survolez une barre pour voir le montant exact.")
    ca_mensuel = df_f.copy(); ca_mensuel["Période"] = ca_mensuel["DATE"].dt.to_period("M").astype(str)
    ca_mb = ca_mensuel.groupby("Période")["MONTANT TTC"].sum().reset_index()
    q33 = ca_mb["MONTANT TTC"].quantile(0.33); q66 = ca_mb["MONTANT TTC"].quantile(0.66)
    ca_mb["Couleur"] = ca_mb["MONTANT TTC"].apply(lambda v: C["vert"] if v>=q66 else (C["orange"] if v>=q33 else C["rouge"]))
    fig = go.Figure(go.Bar(x=ca_mb["Période"], y=ca_mb["MONTANT TTC"],
        marker_color=ca_mb["Couleur"].tolist(),
        text=ca_mb["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
        textposition="outside", textfont=dict(size=9),
        hovertemplate="<b>%{x}</b><br>CA : %{y:,.0f} FCFA<extra></extra>"))
    for lbl, clr in [("🟢 Mois fort",C["vert"]),("🟡 Mois moyen",C["orange"]),("🔴 Mois faible",C["rouge"])]:
        fig.add_trace(go.Scatter(x=[None],y=[None],mode="markers",
            marker=dict(color=clr,size=10,symbol="square"),name=lbl))
    fig = excel_style(fig, 360)
    fig.update_xaxes(tickangle=-40, tickfont=dict(size=8))
    fig.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # COURBES séparées par année
    section("📈  COMPARAISON ANNÉE PAR ANNÉE — Une courbe = une année")
    comment("Si la courbe 2024 est au-dessus de 2023 → nous progressons. Chaque couleur = une année.")
    ca_ann = df_f.groupby(["ANNEE","MOIS"])["MONTANT TTC"].sum().reset_index()
    mois_l = {1:"Jan",2:"Fév",3:"Mar",4:"Avr",5:"Mai",6:"Jun",7:"Jul",8:"Aoû",9:"Sep",10:"Oct",11:"Nov",12:"Déc"}
    ca_ann["Mois_L"] = ca_ann["MOIS"].map(mois_l)
    ann_colors = {2022:C["violet"],2023:C["rose"],2024:C["bleu"]}
    fig2 = go.Figure()
    for an in sorted(ca_ann["ANNEE"].unique()):
        d = ca_ann[ca_ann["ANNEE"]==an].sort_values("MOIS")
        col_an = ann_colors.get(an, C["bleu_fonce"])
        fig2.add_trace(go.Scatter(x=d["Mois_L"], y=d["MONTANT TTC"], mode="lines+markers", name=str(an),
            line=dict(color=col_an, width=3), marker=dict(size=9, color=col_an, line=dict(color="white",width=2)),
            hovertemplate=f"<b>%{{x}} {an}</b><br>CA : %{{y:,.0f}} FCFA<extra></extra>"))
    fig2 = excel_style(fig2, 320)
    fig2.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig2, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # BARRES horizontales CA par segment
    section("🏷️  CA PAR TYPE DE CLIENT — Qui nous rapporte le plus ?")
    comment("La Grande Distribution et les Industriels sont généralement nos premiers acheteurs. "
            "Les distributeurs locaux et l'Export sont des relais de croissance.")
    seg_ca = df_f.groupby("SEGMENT")["MONTANT TTC"].sum().sort_values(ascending=True).reset_index()
    fig3 = go.Figure(go.Bar(x=seg_ca["MONTANT TTC"], y=seg_ca["SEGMENT"], orientation="h",
        marker_color=PIE[:len(seg_ca)],
        text=seg_ca["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
        textposition="outside", textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>CA : %{x:,.0f} FCFA<extra></extra>"))
    fig3 = excel_style(fig3, 300, False)
    fig3.update_xaxes(tickformat=",.0f")
    st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE 3 : NOS PRODUITS
# ══════════════════════════════════════════════════════════════
elif page == "🛍️  Nos Produits":
    sor_agg = df_sor.groupby("Catégorie").agg(CA=("Total","sum"),Qte=("Quantité","sum")).reset_index()
    ent_agg = df_ent.groupby("Catégorie").agg(Achats=("Total","sum")).reset_index()
    perf = sor_agg.merge(ent_agg, on="Catégorie", how="left").fillna(0)
    perf["Marge"] = perf["CA"] - perf["Achats"]
    perf["Taux_m"] = np.where(perf["CA"]>0, perf["Marge"]/perf["CA"]*100, 0).round(1)
    perf = perf.sort_values("CA", ascending=False)
    tm = perf["Marge"].sum()/sor_agg["CA"].sum()*100 if sor_agg["CA"].sum()>0 else 0

    comment("Cette page montre la performance de chaque famille de produits : "
            "Sacs supermarché, Sacs poubelles, Gobelets, Emballages. "
            "La marge = ce qu'on gagne réellement après les matières premières.", "🛍️ Guide de lecture")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("📦","Total Ventes", fmt(sor_agg["CA"].sum()),"Toutes familles","bleu")
    with c2: kpi("🛒","Coût Matières", fmt(ent_agg["Achats"].sum()),"Achats production","orange")
    with c3: kpi("💚","Marge Brute", fmt(perf["Marge"].sum()),"Ventes - Achats","vert")
    with c4: kpi("💹","Taux de Marge", f"{tm:.1f}%","Rentabilité globale","vert" if tm>25 else "orange")
    st.markdown("<br>", unsafe_allow_html=True)

    # Barres CA par famille
    section("📊  VENTES PAR FAMILLE DE PRODUIT")
    comment("Chaque barre = le total des ventes d'une famille. "
            "Les Sacs et Gobelets sont nos produits phares. Un bon CA avec une mauvaise marge = problème à corriger.")
    fig = go.Figure(go.Bar(x=perf["Catégorie"], y=perf["CA"], marker_color=PIE[:len(perf)],
        text=perf["CA"].apply(lambda v: fmt(v,"").strip()),
        textposition="outside", textfont=dict(size=10),
        hovertemplate="<b>%{x}</b><br>Ventes : %{y:,.0f} FCFA<extra></extra>"))
    fig = excel_style(fig, 320, False)
    fig.update_xaxes(tickangle=-15, tickfont=dict(size=10))
    fig.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("💹  MARGE BRUTE PAR FAMILLE")
        comment("La marge = ce que chaque famille nous rapporte réellement. "
                "🟢 Positif = rentable. 🔴 Négatif = on perd de l'argent sur ce produit.")
        perf_s = perf.sort_values("Marge", ascending=True)
        fig2 = go.Figure(go.Bar(x=perf_s["Marge"], y=perf_s["Catégorie"], orientation="h",
            marker_color=[C["vert"] if v>0 else C["rouge"] for v in perf_s["Marge"]],
            text=perf_s["Marge"].apply(lambda v: fmt(v,"").strip()),
            textposition="outside", textfont=dict(size=10),
            hovertemplate="<b>%{y}</b><br>Marge : %{x:,.0f} FCFA<extra></extra>"))
        fig2 = excel_style(fig2, 300, False)
        fig2.update_xaxes(tickformat=",.0f")
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🥧  PART DE CHAQUE FAMILLE DANS LES VENTES")
        comment("Ce graphique rond montre la part (%) de chaque famille dans notre CA total. "
                "Un seul produit dominant = risque si la demande chute. Mieux vaut diversifier.")
        fig3 = go.Figure(go.Pie(labels=perf["Catégorie"], values=perf["CA"], hole=0,
            marker=dict(colors=PIE[:len(perf)], line=dict(color="white",width=2)),
            textinfo="label+percent", textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>%{value:,.0f} FCFA (%{percent})<extra></extra>"))
        fig3.update_layout(paper_bgcolor="#FFFFFF", height=300, margin=dict(t=10,b=30,l=10,r=10),
            legend=dict(orientation="v", x=1.02, y=0.5, font=dict(color=C["texte"],size=10)))
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Courbes séparées par famille
    section("📈  ÉVOLUTION MENSUELLE PAR FAMILLE — Une courbe par gamme de produit")
    comment("Chaque courbe = une famille de produits. Survolez un point pour voir le montant exact. "
            "Une courbe qui monte régulièrement = la demande augmente. Une chute = risque à surveiller.")
    df_sor_m = df_sor.copy()
    df_sor_m["Mois"] = pd.to_datetime(df_sor_m["Date"]).dt.to_period("M").astype(str)
    sm = df_sor_m.groupby(["Mois","Catégorie"])["Total"].sum().reset_index()
    fig4 = go.Figure()
    for j, cat in enumerate(sm["Catégorie"].unique()):
        d = sm[sm["Catégorie"]==cat]; c = PIE[j % len(PIE)]
        fig4.add_trace(go.Scatter(x=d["Mois"], y=d["Total"], name=cat, mode="lines+markers",
            line=dict(color=c, width=2.5), marker=dict(size=6, color=c, line=dict(color="white",width=1.5)),
            hovertemplate=f"<b>{cat}</b><br>%{{x}}<br>%{{y:,.0f}} FCFA<extra></extra>"))
    fig4 = excel_style(fig4, 340)
    fig4.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig4.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE 4 : NOS CLIENTS
# ══════════════════════════════════════════════════════════════
elif page == "👥  Nos Clients":
    cl_stats = df_f.groupby(["CLIENT","SEGMENT","ZONE"]).agg(
        CA=("MONTANT TTC","sum"), Nb_Cmd=("MONTANT TTC","count"), Panier=("MONTANT TTC","mean"),
        CA_P=("MONTANT TTC", lambda x: x[df_f.loc[x.index,"ETAT DE PAIEMENT"]=="Payée"].sum()),
        CA_I=("MONTANT TTC", lambda x: x[df_f.loc[x.index,"ETAT DE PAIEMENT"]=="Impayée"].sum()),
    ).reset_index()
    cl_stats["Taux_R"] = np.where(cl_stats["CA"]>0, cl_stats["CA_P"]/cl_stats["CA"]*100, 0).round(1)
    cl_stats = cl_stats.sort_values("CA", ascending=False)

    comment(f"Nous avons <b>{len(cl_stats)}</b> clients actifs. Un client fidèle coûte 5x moins cher "
            f"à garder qu'à en trouver un nouveau. Concentrons nos efforts sur les meilleurs.", "👥 Pourquoi analyser nos clients ?")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("👥","Total Clients", str(len(cl_stats)),"Au moins 1 commande","bleu")
    with c2: kpi("⭐","Clients Fidèles", str((cl_stats["Nb_Cmd"]>=5).sum()),"5 commandes ou +","vert")
    with c3: kpi("⚠️","Clients Impayés", str((cl_stats["CA_I"]>0).sum()),"À relancer","rouge")
    with c4: kpi("🏆","CA Moyen/Client", fmt(cl_stats["CA"].mean()),"Par client actif","violet")
    st.markdown("<br>", unsafe_allow_html=True)

    # Top 10 clients
    section("🏆  TOP 10 CLIENTS — Classement par chiffre d'affaires")
    comment("Nos 10 meilleurs clients. Une perte d'un de ces clients aurait un impact fort sur notre CA. "
            "Visitez-les régulièrement et soignez la relation.")
    top10 = cl_stats.head(10).sort_values("CA", ascending=True)
    fig = go.Figure(go.Bar(x=top10["CA"], y=top10["CLIENT"], orientation="h", marker_color=C["bleu"],
        text=top10["CA"].apply(lambda v: fmt(v,"").strip()),
        textposition="outside", textfont=dict(size=10),
        hovertemplate="<b>%{y}</b><br>CA : %{x:,.0f} FCFA<extra></extra>"))
    fig = excel_style(fig, 360, False)
    fig.update_xaxes(tickformat=",.0f")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("🗂️  CA PAR TYPE DE CLIENT")
        comment("Répartition des ventes par catégorie de clients. "
                "Trop dépendre d'un seul type = risque. Mieux vaut diversifier.")
        seg_ca = cl_stats.groupby("SEGMENT")["CA"].sum().reset_index()
        fig2 = go.Figure(go.Pie(labels=seg_ca["SEGMENT"], values=seg_ca["CA"], hole=0.60,
            marker=dict(colors=PIE[:len(seg_ca)], line=dict(color="white",width=2)),
            textinfo="label+percent", textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>%{value:,.0f} FCFA (%{percent})<extra></extra>"))
        fig2.update_layout(paper_bgcolor="#FFFFFF", height=320, margin=dict(t=10,b=30,l=10,r=10),
            legend=dict(orientation="h", y=-0.08, xanchor="center", x=0.5, font=dict(color=C["texte"],size=10)))
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🗺️  CA PAYÉ vs IMPAYÉ PAR ZONE")
        comment("Bleu = encaissé. Rose = factures non payées. "
                "Une zone avec beaucoup de rose = relancer les clients de cette région en priorité.")
        zone_ca = cl_stats.groupby("ZONE")[["CA_P","CA_I"]].sum().reset_index()
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(x=zone_ca["ZONE"], y=zone_ca["CA_P"], name="Payé ✅", marker_color=C["bleu"]))
        fig3.add_trace(go.Bar(x=zone_ca["ZONE"], y=zone_ca["CA_I"], name="Impayé ⏳", marker_color=C["rose"]))
        fig3.update_layout(barmode="stack")
        fig3 = excel_style(fig3, 320)
        fig3.update_xaxes(tickangle=-15, tickfont=dict(size=10))
        fig3.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig3, use_container_width=True)

    section("📋  FICHE COMPLÈTE NOS CLIENTS")
    disp = cl_stats[["CLIENT","SEGMENT","ZONE","CA","Nb_Cmd","Panier","CA_P","CA_I","Taux_R"]].copy()
    for col_n in ["CA","CA_P","CA_I","Panier"]:
        disp[col_n] = disp[col_n].apply(lambda x: f"{x:,.0f} FCFA")
    disp["Taux_R"] = cl_stats["Taux_R"].apply(lambda x: f"{x:.1f}%")
    disp = disp.rename(columns={"CLIENT":"Client","SEGMENT":"Segment","ZONE":"Zone",
        "CA":"CA Total","Nb_Cmd":"Commandes","Panier":"Panier Moy.","CA_P":"Payé","CA_I":"Impayé","Taux_R":"% Payé"})
    st.dataframe(disp, use_container_width=True, height=350)

# ══════════════════════════════════════════════════════════════
# PAGE 5 : STOCKS & INVENTAIRE
# ══════════════════════════════════════════════════════════════
elif page == "📦  Stocks & Inventaire":
    val_stock = df_inv["Valeur"].sum()
    nb_rupt   = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib   = (df_inv["Statut"]=="Stock faible").sum()
    nb_norm   = (df_inv["Statut"]=="Stock normal").sum()

    comment(f"Notre stock total vaut <b>{fmt(val_stock)}</b>. "
            f"<b>{nb_rupt}</b> produit(s) en rupture totale et <b>{nb_faib}</b> approchent de la limite. "
            f"Un bon stock = on ne rate aucune commande client.", "📦 Pourquoi surveiller les stocks ?")

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","Valeur du Stock", fmt(val_stock), f"{len(df_inv)} références","bleu")
    with c2: kpi("✅","Stock Normal", str(nb_norm),"Niveau satisfaisant","vert")
    with c3: kpi("🟡","Stock Faible", str(nb_faib),"Commander bientôt","orange")
    with c4: kpi("🔴","Rupture Totale", str(nb_rupt),"Commander d'urgence","rouge")
    with c5:
        rot = df_sor["Quantité"].sum()/max(df_inv["Stock final"].sum(),1)
        kpi("🔄","Rotation Stock", f"{rot:.1f}×","Fréquence renouvellement","violet")
    st.markdown("<br>", unsafe_allow_html=True)

    # Barres stock par catégorie colorées
    section("📊  STOCK DISPONIBLE PAR FAMILLE")
    comment("🟢 Vert = stock suffisant &nbsp; 🟡 Orange = à surveiller &nbsp; 🔴 Rouge = rupture imminente. "
            "La ligne pointillée = le seuil minimum acceptable.")
    cat_s = df_inv.groupby("categorie").agg(Stock=("Stock final","sum"), Seuil=("seuil","sum")).reset_index()
    clr_s = [C["vert"] if row["Stock"]>=row["Seuil"] else (C["orange"] if row["Stock"]>0 else C["rouge"])
              for _,row in cat_s.iterrows()]
    fig = go.Figure(go.Bar(x=cat_s["categorie"], y=cat_s["Stock"], marker_color=clr_s,
        text=cat_s["Stock"].apply(lambda v: f"{v:,} u."),
        textposition="outside", textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>Stock : %{y:,} unités<extra></extra>"))
    fig = excel_style(fig, 320, False)
    fig.update_xaxes(tickangle=-15, tickfont=dict(size=10))
    fig.update_yaxes(tickformat=",")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("💰  VALEUR DU STOCK PAR FAMILLE")
        comment("Combien vaut en FCFA chaque famille en stock. "
                "Trop de stock coûte de l'argent (entrepôt, risque) — pas assez = ventes ratées.")
        cat_v = df_inv.groupby("categorie")["Valeur"].sum().reset_index()
        fig2 = go.Figure(go.Pie(labels=cat_v["categorie"], values=cat_v["Valeur"], hole=0.55,
            marker=dict(colors=PIE[:len(cat_v)], line=dict(color="white",width=2)),
            textinfo="label+percent", textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>%{value:,.0f} FCFA (%{percent})<extra></extra>"))
        fig2.update_layout(paper_bgcolor="#FFFFFF", height=300, margin=dict(t=10,b=30,l=10,r=10),
            legend=dict(orientation="v", x=1.02, y=0.5, font=dict(color=C["texte"],size=10)))
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🚦  ÉTAT DE SANTÉ DU STOCK")
        comment("🟢 Normal : tout va bien. 🟡 Faible : commander prochainement. "
                "🔴 Non disponible : rupture totale, risque de perdre des ventes.")
        statuts = ["Stock normal","Stock faible","Non disponible"]
        nbs = [(df_inv["Statut"]==s).sum() for s in statuts]
        fig3 = go.Figure(go.Bar(x=nbs, y=statuts, orientation="h",
            marker_color=[C["vert"],C["orange"],C["rouge"]],
            text=[f"{n} référence(s)" for n in nbs],
            textposition="outside", textfont=dict(size=12),
            hovertemplate="<b>%{y}</b> : %{x} référence(s)<extra></extra>"))
        fig3 = excel_style(fig3, 300, False)
        fig3.update_xaxes(range=[0, max(nbs)*1.4])
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Courbe achats mensuels
    section("📥  ACHATS DE MATIÈRES — Évolution mensuelle")
    comment("Combien nous avons dépensé chaque mois pour acheter nos matières premières. "
            "Des dépenses élevées sans ventes correspondantes = attention à la trésorerie.")
    df_em = df_ent.copy(); df_em["Mois"] = pd.to_datetime(df_em["Date"]).dt.to_period("M").astype(str)
    em_mois = df_em.groupby("Mois")["Total"].sum().reset_index()
    fig4 = go.Figure(go.Scatter(x=em_mois["Mois"], y=em_mois["Total"], mode="lines+markers",
        line=dict(color=C["violet"], width=3),
        marker=dict(size=8, color=C["violet"], line=dict(color="white",width=2)),
        fill="tozeroy", fillcolor="rgba(81,69,158,0.10)",
        hovertemplate="<b>%{x}</b><br>Achats : %{y:,.0f} FCFA<extra></extra>"))
    fig4 = excel_style(fig4, 280, False)
    fig4.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig4.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig4, use_container_width=True)

    section("📋  INVENTAIRE COMPLET")
    inv_d = df_inv[["ref","designation","categorie","seuil","Stock initial","Entrées","Sorties","Stock final","Valeur","Statut"]].copy()
    inv_d["Valeur"] = inv_d["Valeur"].apply(lambda x: f"{x:,.0f} FCFA")
    inv_d = inv_d.rename(columns={"ref":"Réf.","designation":"Désignation","categorie":"Catégorie","seuil":"Seuil"})
    st.dataframe(inv_d, use_container_width=True, height=340)


# ══════════════════════════════════════════════════════════════
# PAGE 6 : PRODUCTION
# ══════════════════════════════════════════════════════════════
elif page == "🏭  Production":
    rend_moy = df_prod["Taux rendement"].mean()
    rebut_moy= df_prod["Taux rebut"].mean()
    prod_tot = df_prod["Production réelle"].sum()
    reb_tot  = df_prod["Rebuts"].sum()

    comment(f"Nos machines ont produit <b>{prod_tot:,.0f} unités</b> avec un rendement de <b>{rend_moy:.1f}%</b>. "
            f"Taux de rebut (produits ratés) : <b>{rebut_moy:.2f}%</b>. Objectif : atteindre 85%.", "🏭 Comment lire la production ?")

    c1,c2,c3,c4 = st.columns(4)
    with c1: kpi("⚙️","Production Totale", f"{prod_tot:,.0f}","unités fabriquées","bleu")
    with c2: kpi("📈","Rendement Moyen", f"{rend_moy:.1f}%","Objectif : 85%","vert" if rend_moy>=85 else "orange")
    with c3: kpi("🗑️","Taux Rebut Moyen", f"{rebut_moy:.2f}%","Objectif < 3%","rouge" if rebut_moy>=3 else "vert")
    with c4: kpi("🔧","Unités Perdues", f"{reb_tot:,.0f}","Rebuts cumulés — coût caché","orange")
    st.markdown("<br>", unsafe_allow_html=True)

    # Courbe production mensuelle globale
    section("📈  PRODUCTION MENSUELLE — Évolution du volume produit")
    comment("Combien d'unités nos machines ont fabriqué chaque mois. "
            "Les pics = forte demande. Les creux = panne, maintenance ou manque de commandes.")
    pm = df_prod.groupby("Mois_Label").agg(Réelle=("Production réelle","sum")).reset_index()
    try:
        pm["sk"] = pd.to_datetime(pm["Mois_Label"], format="%b %Y"); pm = pm.sort_values("sk")
    except: pass
    fig = go.Figure(go.Scatter(x=pm["Mois_Label"], y=pm["Réelle"], mode="lines+markers",
        line=dict(color=C["bleu"], width=3), marker=dict(size=8, color=C["bleu"], line=dict(color="white",width=2)),
        fill="tozeroy", fillcolor="rgba(97,114,243,0.10)",
        hovertemplate="<b>%{x}</b><br>Production : %{y:,.0f} unités<extra></extra>"))
    fig = excel_style(fig, 300, False)
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig.update_yaxes(tickformat=",")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("⚙️  RENDEMENT PAR MACHINE")
        comment("🟢 Au-dessus de 85% = excellent. 🟡 75-85% = acceptable. 🔴 Sous 75% = maintenance nécessaire. "
                "La ligne rouge = notre objectif.")
        rm = df_prod.groupby("Machine")["Taux rendement"].mean().reset_index().sort_values("Taux rendement")
        clr_m = [C["vert"] if v>=85 else C["orange"] if v>=75 else C["rouge"] for v in rm["Taux rendement"]]
        fig2 = go.Figure(go.Bar(x=rm["Taux rendement"], y=rm["Machine"], orientation="h",
            marker_color=clr_m, text=rm["Taux rendement"].apply(lambda v: f"{v:.1f}%"),
            textposition="outside", textfont=dict(size=12),
            hovertemplate="<b>%{y}</b><br>Rendement : %{x:.1f}%<extra></extra>"))
        fig2.add_vline(x=85, line_dash="dot", line_color=C["rouge"], line_width=2,
            annotation_text="Objectif 85%", annotation_font=dict(color=C["rouge"],size=11))
        fig2 = excel_style(fig2, 300, False)
        fig2.update_xaxes(ticksuffix="%", range=[60,108])
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🗑️  REBUTS PAR MACHINE")
        comment("Les rebuts = produits fabriqués mais inutilisables. "
                "Moins il y en a, mieux c'est. Une machine avec beaucoup de rebuts doit être révisée.")
        rb = df_prod.groupby("Machine")["Rebuts"].sum().reset_index().sort_values("Rebuts", ascending=True)
        fig3 = go.Figure(go.Bar(x=rb["Rebuts"], y=rb["Machine"], orientation="h",
            marker_color=C["rose"], text=rb["Rebuts"].apply(lambda v: f"{v:,.0f}"),
            textposition="outside", textfont=dict(size=11),
            hovertemplate="<b>%{y}</b><br>Rebuts : %{x:,.0f} unités<extra></extra>"))
        fig3 = excel_style(fig3, 300, False)
        fig3.update_xaxes(tickformat=",")
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Courbe taux rebut mensuel
    section("📉  ÉVOLUTION DU TAUX DE REBUT — Est-ce qu'on s'améliore ?")
    comment("Une courbe qui descend = on s'améliore. Une courbe qui monte = problème de qualité. "
            "L'objectif est de maintenir ce taux sous 3% (ligne pointillée).")
    rb_m = df_prod.groupby("Mois_Label")["Taux rebut"].mean().reset_index()
    try:
        rb_m["sk"] = pd.to_datetime(rb_m["Mois_Label"], format="%b %Y"); rb_m = rb_m.sort_values("sk")
    except: pass
    fig4 = go.Figure(go.Scatter(x=rb_m["Mois_Label"], y=rb_m["Taux rebut"], mode="lines+markers",
        line=dict(color=C["rose"], width=3), marker=dict(size=7, color=C["rose"], line=dict(color="white",width=2)),
        fill="tozeroy", fillcolor="rgba(255,107,167,0.10)",
        hovertemplate="<b>%{x}</b><br>Taux rebut : %{y:.2f}%<extra></extra>"))
    fig4.add_hline(y=3, line_dash="dot", line_color=C["rouge"], line_width=1.5,
        annotation_text="Seuil max 3%", annotation_font=dict(color=C["rouge"],size=10))
    fig4 = excel_style(fig4, 280, False)
    fig4.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig4.update_yaxes(ticksuffix="%")
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE 7 : PAIEMENTS
# ══════════════════════════════════════════════════════════════
elif page == "💳  Paiements":
    df_p = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]
    df_i = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]
    taux  = len(df_p)/len(df_f)*100 if len(df_f) else 0
    mode_top = df_p["MODE DE PAIEMENT"].value_counts().idxmax() if len(df_p)>0 else "N/A"
    mode_colors = [C["bleu"],C["rose"],C["violet"],C["bleu_clair"],C["bleu_fonce"]]

    comment(f"Sur <b>{len(df_f)}</b> factures : <b>{len(df_p)}</b> payées ({taux:.1f}%) et "
            f"<b>{len(df_i)}</b> impayées. Mode dominant : <b>{mode_top}</b>. "
            f"Un bon taux de recouvrement = moins de risques financiers.", "💳 Comprendre nos paiements")

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: kpi("💰","Total Encaissé", fmt(df_p["MONTANT TTC"].sum()),"Reçus","vert")
    with c2: kpi("⏳","Total Impayé", fmt(df_i["MONTANT TTC"].sum()), f"{len(df_i)} fact.","rouge")
    with c3: kpi("📊","Taux Recouvrement", f"{taux:.1f}%","% factures payées","bleu")
    with c4: kpi("🧾","Montant Moy. Payé", fmt(df_p["MONTANT TTC"].mean()),"Par facture réglée","violet")
    with c5: kpi("🏆","Mode Dominant", mode_top.split()[0],"Le + utilisé","orange")
    st.markdown("<br>", unsafe_allow_html=True)

    # Courbe taux recouvrement mensuel
    section("📈  TAUX DE RECOUVREMENT MENSUEL — Récupère-t-on bien notre argent ?")
    comment("Au-dessus de 80% (ligne verte) = bonne santé financière. "
            "En dessous = relancer les clients. Les points rouges = mois en dessous de l'objectif.")
    df_fm = df_f.copy(); df_fm["Période"] = df_fm["DATE"].dt.to_period("M").astype(str)
    rec_m = df_fm.groupby("Période").apply(
        lambda g: g[g["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()/g["MONTANT TTC"].sum()*100
        if g["MONTANT TTC"].sum()>0 else 0).reset_index()
    rec_m.columns = ["Période","Taux"]
    fig = go.Figure(go.Scatter(x=rec_m["Période"], y=rec_m["Taux"], mode="lines+markers",
        line=dict(color=C["bleu"], width=3),
        marker=dict(size=9, color=[C["vert"] if v>=80 else C["rouge"] for v in rec_m["Taux"]],
                    line=dict(color="white",width=2)),
        fill="tozeroy", fillcolor="rgba(97,114,243,0.08)",
        hovertemplate="<b>%{x}</b><br>Recouvrement : %{y:.1f}%<extra></extra>"))
    fig.add_hline(y=80, line_dash="dot", line_color=C["vert"], line_width=2,
        annotation_text="Objectif 80% ✅", annotation_font=dict(color=C["vert"],size=11),
        annotation_position="top right")
    fig = excel_style(fig, 300, False)
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig.update_yaxes(ticksuffix="%", range=[40,110])
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        section("🥧  COMMENT NOS CLIENTS PAIENT-ILS ?")
        comment("Le virement bancaire est le plus sécurisé. "
                "Un fort pourcentage d'espèces peut rendre le suivi plus difficile.")
        md = df_p.groupby("MODE DE PAIEMENT")["MONTANT TTC"].sum().reset_index().sort_values("MONTANT TTC",ascending=False)
        fig2 = go.Figure(go.Pie(labels=md["MODE DE PAIEMENT"], values=md["MONTANT TTC"], hole=0.60,
            marker=dict(colors=mode_colors[:len(md)], line=dict(color="white",width=2)),
            textinfo="label+percent", textfont=dict(size=11),
            hovertemplate="<b>%{label}</b><br>%{value:,.0f} FCFA (%{percent})<extra></extra>"))
        fig2.update_layout(paper_bgcolor="#FFFFFF", height=300, margin=dict(t=10,b=30,l=10,r=10),
            legend=dict(orientation="h", y=-0.08, xanchor="center", x=0.5, font=dict(color=C["texte"],size=10)))
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        section("🔴  CLIENTS LES PLUS EN RETARD")
        comment("Ces clients ont les montants impayés les plus élevés. "
                "À contacter en priorité. Plus une facture est vieille, plus elle risque d'être perdue.")
        top_i = df_i.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=True).tail(8).reset_index()
        fig3 = go.Figure(go.Bar(x=top_i["MONTANT TTC"], y=top_i["CLIENT"], orientation="h",
            marker_color=C["rouge"], text=top_i["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
            textposition="outside", textfont=dict(size=10),
            hovertemplate="<b>%{y}</b><br>Impayé : %{x:,.0f} FCFA<extra></extra>"))
        fig3 = excel_style(fig3, 300, False)
        fig3.update_xaxes(tickformat=",.0f")
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Barres groupées (non empilées) par mode et mois
    section("📅  ENCAISSEMENTS PAR MODE DE PAIEMENT — Barres séparées par mois")
    comment("Chaque couleur = un mode de paiement. Les barres sont séparées pour faciliter la lecture. "
            "Permettez-vous de voir si un mode de paiement progresse ou régresse dans le temps.")
    df_pm = df_p.copy(); df_pm["Période"] = df_pm["DATE"].dt.to_period("M").astype(str)
    mm = df_pm.groupby(["Période","MODE DE PAIEMENT"])["MONTANT TTC"].sum().reset_index()
    fig4 = go.Figure()
    for j, m in enumerate(mm["MODE DE PAIEMENT"].unique()):
        d = mm[mm["MODE DE PAIEMENT"]==m]
        fig4.add_trace(go.Bar(x=d["Période"], y=d["MONTANT TTC"], name=m,
            marker_color=mode_colors[j % len(mode_colors)],
            hovertemplate=f"<b>{m}</b><br>%{{x}}<br>%{{y:,.0f}} FCFA<extra></extra>"))
    fig4.update_layout(barmode="group")
    fig4 = excel_style(fig4, 320)
    fig4.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig4.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PAGE 8 : PRÉVISIONS & ANTICIPATIONS
# ══════════════════════════════════════════════════════════════
elif page == "🔮  Prévisions & Anticipations":

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1E2A4A,#2D3A6B);border-radius:14px;
                padding:22px 28px;margin-bottom:20px;border:1px solid rgba(97,114,243,0.3);">
        <div style="font-size:1.15rem;font-weight:800;color:white;margin-bottom:4px;">
            🔮 Prévisions & Anticipations — MULTIPACK SA
        </div>
        <div style="font-size:0.8rem;color:rgba(255,255,255,0.65);">
            Basé sur l'analyse des tendances historiques 2022–2024 · Modèle de régression polynomiale
        </div>
    </div>""", unsafe_allow_html=True)

    prevision_comment(
        "Ces prévisions sont calculées automatiquement à partir de l'historique de vos données. "
        "Elles vous donnent une <b>estimation des 6 prochains mois</b>. "
        "Ce ne sont pas des certitudes, mais des signaux pour mieux <b>anticiper et planifier</b>. "
        "La zone ombrée représente la fourchette haute et basse probable.",
        "🔮 Comment lire ces prévisions ?"
    )

    n_mois_prev = st.slider("Nombre de mois à prévoir", min_value=3, max_value=12, value=6, step=1)

    # ──────────────────────────────────────────────
    # 1. PRÉVISION CA GLOBAL
    # ──────────────────────────────────────────────
    section("📈  PRÉVISION DU CHIFFRE D'AFFAIRES — 6 prochains mois")
    prevision_comment(
        f"En prolongeant votre tendance actuelle, voici ce que devrait être votre CA sur les {n_mois_prev} prochains mois. "
        f"La zone bleue claire = intervalle de confiance (fourchette basse–haute). "
        f"Si la courbe monte = la tendance est favorable. Si elle descend = prendre des mesures dès maintenant."
    )

    df_f_all = df_fact.copy()
    df_f_all["Période"] = df_f_all["DATE"].dt.to_period("M").astype(str)
    serie_hist = df_f_all.groupby("Période")["MONTANT TTC"].sum()
    prev_df = prevoir_ca(serie_hist, n_mois_prev)

    hist_x = serie_hist.index.tolist()
    hist_y = serie_hist.values.tolist()

    fig = go.Figure()
    # Historique
    fig.add_trace(go.Scatter(x=hist_x, y=hist_y, name="CA Historique (réel)",
        mode="lines+markers", line=dict(color=C["bleu"], width=2.5),
        marker=dict(size=5, color=C["bleu"]), fill="tozeroy",
        fillcolor="rgba(97,114,243,0.08)",
        hovertemplate="<b>%{x}</b><br>CA réel : %{y:,.0f} FCFA<extra></extra>"))
    # Zone de confiance
    fig.add_trace(go.Scatter(x=list(prev_df["Période"])+list(prev_df["Période"])[::-1],
        y=list(prev_df["Haut"])+list(prev_df["Bas"])[::-1],
        fill="toself", fillcolor="rgba(247,179,49,0.15)",
        line=dict(color="rgba(0,0,0,0)"), showlegend=True, name="Fourchette probable",
        hoverinfo="skip"))
    # Courbe prévision
    fig.add_trace(go.Scatter(x=prev_df["Période"], y=prev_df["CA_Prévu"],
        name="Prévision CA", mode="lines+markers",
        line=dict(color=C["orange"], width=3, dash="dash"),
        marker=dict(size=9, color=C["orange"], symbol="diamond", line=dict(color="white",width=2)),
        hovertemplate="<b>Prévision %{x}</b><br>CA estimé : %{y:,.0f} FCFA<extra></extra>"))
    # Séparateur historique / prévision
    if hist_x:
        fig.add_vline(x=hist_x[-1], line_dash="dot", line_color=C["muted"], line_width=1,
            annotation_text="Aujourd'hui", annotation_font=dict(color=C["muted"],size=10))
    fig = excel_style(fig, 380)
    fig.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig, use_container_width=True)

    # KPIs prévision CA
    ca_prev_total = prev_df["CA_Prévu"].sum()
    ca_prev_bas   = prev_df["Bas"].sum()
    ca_prev_haut  = prev_df["Haut"].sum()
    ca_moy_hist   = serie_hist.mean()
    evol_prev = (prev_df["CA_Prévu"].mean() - ca_moy_hist) / ca_moy_hist * 100 if ca_moy_hist else 0

    st.markdown("<br>", unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1: prev_kpi("🎯","CA Prévu Total", fmt(ca_prev_total), f"Sur {n_mois_prev} mois")
    with c2: prev_kpi("📉","Scénario Pessimiste", fmt(ca_prev_bas), "Fourchette basse")
    with c3: prev_kpi("📈","Scénario Optimiste", fmt(ca_prev_haut), "Fourchette haute")
    with c4: prev_kpi("📊","Évolution Prévue", f"{evol_prev:+.1f}%", "vs moyenne historique")

    st.markdown("<br>", unsafe_allow_html=True)

    # Tableau détaillé des prévisions CA
    with st.expander("📋 Voir le détail mois par mois des prévisions CA"):
        prev_disp = prev_df.copy()
        prev_disp["CA_Prévu"] = prev_disp["CA_Prévu"].apply(lambda v: f"{v:,.0f} FCFA")
        prev_disp["Bas"]      = prev_disp["Bas"].apply(lambda v: f"{v:,.0f} FCFA")
        prev_disp["Haut"]     = prev_disp["Haut"].apply(lambda v: f"{v:,.0f} FCFA")
        prev_disp = prev_disp.rename(columns={"Période":"Mois","CA_Prévu":"CA Estimé","Bas":"Scénario Bas","Haut":"Scénario Haut"})
        st.dataframe(prev_disp, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ──────────────────────────────────────────────
    # 2. PRÉVISION PAR SEGMENT
    # ──────────────────────────────────────────────
    section("🏷️  PRÉVISION PAR SEGMENT CLIENT — Qui va progresser ?")
    prevision_comment(
        "Cette analyse prédit l'évolution de chaque segment client pour les prochains mois. "
        "Un segment en hausse = opportunité commerciale à saisir. "
        "Un segment en baisse = risque à anticiper et à corriger."
    )

    df_fa = df_fact.copy()
    df_fa["Période"] = df_fa["DATE"].dt.to_period("M").astype(str)
    segments = df_fa["SEGMENT"].unique()
    seg_colors_map = dict(zip(segments, PIE))

    fig2 = go.Figure()
    seg_prev_summary = []
    for seg in segments:
        serie_seg = df_fa[df_fa["SEGMENT"]==seg].groupby("Période")["MONTANT TTC"].sum()
        if len(serie_seg) < 4:
            continue
        p = prevoir_ca(serie_seg, n_mois_prev)
        c = seg_colors_map.get(seg, C["bleu"])
        fig2.add_trace(go.Scatter(x=p["Période"], y=p["CA_Prévu"], name=seg,
            mode="lines+markers", line=dict(color=c, width=2.5, dash="dash"),
            marker=dict(size=8, symbol="diamond", color=c, line=dict(color="white",width=1.5)),
            hovertemplate=f"<b>{seg}</b><br>%{{x}}<br>Prévision : %{{y:,.0f}} FCFA<extra></extra>"))
        moy_hist = serie_seg.mean()
        moy_prev = p["CA_Prévu"].mean()
        seg_prev_summary.append({"Segment": seg, "Moy. Hist.": moy_hist, "Moy. Prévue": moy_prev,
                                   "Tendance": (moy_prev-moy_hist)/moy_hist*100 if moy_hist else 0})

    fig2 = excel_style(fig2, 340)
    fig2.update_xaxes(tickangle=-30, tickfont=dict(size=9))
    fig2.update_yaxes(tickformat=",.0f", ticksuffix=" F")
    st.plotly_chart(fig2, use_container_width=True)

    # Résumé tendances segments
    if seg_prev_summary:
        st.markdown("<br>", unsafe_allow_html=True)
        cols_seg = st.columns(len(seg_prev_summary))
        for i, s in enumerate(seg_prev_summary):
            t = s["Tendance"]
            icon  = "📈" if t>5 else ("📉" if t<-5 else "➡️")
            color = "vert" if t>5 else ("rouge" if t<-5 else "orange")
            with cols_seg[i]:
                kpi(icon, s["Segment"][:20], f"{t:+.1f}%",
                    f"Moy. prévue : {fmt(s['Moy. Prévue'],'').strip()}", color)

    st.markdown("<br>", unsafe_allow_html=True)

    # ──────────────────────────────────────────────
    # 3. PRÉVISION STOCKS — RUPTURES IMMINENTES
    # ──────────────────────────────────────────────
    section("📦  ANTICIPATION DES RUPTURES DE STOCK — Quand faudra-t-il commander ?")
    prevision_comment(
        "Pour chaque produit, nous calculons combien de mois de stock il reste "
        "en fonction du rythme moyen des ventes. "
        "🔴 Moins de 2 mois = commander maintenant. 🟡 2–4 mois = planifier bientôt. 🟢 Plus de 4 mois = situation confortable."
    )

    stock_alert = []
    for _, row in df_inv.iterrows():
        sf  = row["Stock final"]
        sor = row.get("Sorties_moy_mois", 1)
        seuil = row["seuil"]
        if sor <= 0: sor = 1
        mois_restants = (sf - seuil) / sor if sor > 0 else 99
        stock_alert.append({
            "Produit":        row["designation"][:35],
            "Catégorie":      row["categorie"],
            "Stock actuel":   int(sf),
            "Sorties/mois":   round(float(sor), 1),
            "Mois restants":  max(0, round(mois_restants, 1)),
            "Statut actuel":  row["Statut"],
        })

    df_alert = pd.DataFrame(stock_alert).sort_values("Mois restants")

    # Graphique mois restants
    df_graph = df_alert[df_alert["Mois restants"] < 10].head(15).sort_values("Mois restants", ascending=True)
    clr_alert = [C["rouge"] if v<2 else (C["orange"] if v<4 else C["vert"]) for v in df_graph["Mois restants"]]
    fig3 = go.Figure(go.Bar(
        x=df_graph["Mois restants"], y=df_graph["Produit"],
        orientation="h", marker_color=clr_alert,
        text=df_graph["Mois restants"].apply(lambda v: f"{v:.1f} mois"),
        textposition="outside", textfont=dict(size=11),
        hovertemplate="<b>%{y}</b><br>Stock pour : %{x:.1f} mois<extra></extra>"))
    fig3.add_vline(x=2, line_dash="dot", line_color=C["rouge"], line_width=1.5,
        annotation_text="⚠️ Commander maintenant", annotation_font=dict(color=C["rouge"],size=10))
    fig3.add_vline(x=4, line_dash="dot", line_color=C["orange"], line_width=1.5,
        annotation_text="📅 Planifier", annotation_font=dict(color=C["orange"],size=10))
    fig3 = excel_style(fig3, 420, False)
    fig3.update_xaxes(ticksuffix=" mois", range=[0, max(df_graph["Mois restants"].max()*1.3, 5)])
    st.plotly_chart(fig3, use_container_width=True)

    # Alertes ruptures imminentes
    urgents = df_alert[df_alert["Mois restants"] < 2]
    attentions = df_alert[(df_alert["Mois restants"] >= 2) & (df_alert["Mois restants"] < 4)]
    if len(urgents) > 0:
        for _, r in urgents.head(5).iterrows():
            st.markdown(f"""<div class="alert-r">🔴<div>
                <b>COMMANDER MAINTENANT — {r['Produit']}</b><br>
                <small>Stock restant : {r['Stock actuel']} u. · Consommation : ~{r['Sorties/mois']:.0f} u./mois
                → Stock épuisé dans <b>{r['Mois restants']:.1f} mois</b></small>
            </div></div>""", unsafe_allow_html=True)
    if len(attentions) > 0:
        for _, r in attentions.head(4).iterrows():
            st.markdown(f"""<div class="alert-y">📅<div>
                <b>PLANIFIER BIENTÔT — {r['Produit']}</b><br>
                <small>Stock pour <b>{r['Mois restants']:.1f} mois</b>
                · Consommation : ~{r['Sorties/mois']:.0f} u./mois</small>
            </div></div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ──────────────────────────────────────────────
    # 4. PRÉVISION RECOUVREMENT & TRÉSORERIE
    # ──────────────────────────────────────────────
    section("💰  ANTICIPATION DE TRÉSORERIE — Ce qu'on devrait encaisser")
    prevision_comment(
        "En appliquant notre taux de recouvrement moyen au CA prévu, "
        "voici ce que nous pouvons anticiper comme entrées de trésorerie. "
        "Ce chiffre est utile pour planifier les dépenses, les investissements et les achats de matières."
    )

    taux_rec_hist = df_fact[df_fact["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum() / df_fact["MONTANT TTC"].sum()
    prev_encaisse = prev_df["CA_Prévu"] * taux_rec_hist
    prev_impaye   = prev_df["CA_Prévu"] * (1 - taux_rec_hist)

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=prev_df["Période"], y=prev_encaisse,
            name=f"Encaissements attendus ({taux_rec_hist*100:.0f}%)",
            marker_color=C["vert"], opacity=0.85,
            hovertemplate="<b>%{x}</b><br>Encaissé prévu : %{y:,.0f} FCFA<extra></extra>"))
        fig4.add_trace(go.Bar(x=prev_df["Période"], y=prev_impaye,
            name="Risque impayé estimé", marker_color=C["rose"], opacity=0.75,
            hovertemplate="<b>%{x}</b><br>Risque impayé : %{y:,.0f} FCFA<extra></extra>"))
        fig4.update_layout(barmode="stack")
        fig4 = excel_style(fig4, 300)
        fig4.update_xaxes(tickangle=-30, tickfont=dict(size=9))
        fig4.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig4, use_container_width=True)

    with col_t2:
        enc_total = prev_encaisse.sum()
        imp_total = prev_impaye.sum()
        st.markdown("<br><br>", unsafe_allow_html=True)
        prev_kpi("💚","Encaissements Prévus", fmt(enc_total), f"Sur {n_mois_prev} mois (taux {taux_rec_hist*100:.0f}%)")
        st.markdown("<br>", unsafe_allow_html=True)
        prev_kpi("⚠️","Risque Impayé Estimé", fmt(imp_total), "À surveiller et relancer")
        st.markdown("<br>", unsafe_allow_html=True)
        prevision_comment(
            f"Sur les {n_mois_prev} prochains mois, nous anticipons <b>{fmt(enc_total)}</b> d'encaissements. "
            f"Pour améliorer ce chiffre, il faut réduire les impayés : "
            f"relances régulières, conditions de paiement claires, acomptes sur grandes commandes.",
            "💡 Comment améliorer la trésorerie ?"
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ──────────────────────────────────────────────
    # 5. PRÉVISION PRODUCTION
    # ──────────────────────────────────────────────
    section("🏭  PRÉVISION DE PRODUCTION — Planifier la capacité usine")
    prevision_comment(
        "En prolongeant les tendances de production, voici ce que nos machines devraient produire. "
        "Si la prévision dépasse notre capacité actuelle → il faut prévoir des heures supplémentaires ou de nouveaux équipements. "
        "Si elle est en baisse → surveiller les carnets de commandes."
    )

    prod_mois = df_prod.groupby("Mois_Label")["Production réelle"].sum().reset_index()
    try:
        prod_mois["sk"] = pd.to_datetime(prod_mois["Mois_Label"], format="%b %Y")
        prod_mois = prod_mois.sort_values("sk")
    except: pass
    serie_prod = pd.Series(prod_mois["Production réelle"].values, index=prod_mois["Mois_Label"])
    prev_prod = prevoir_ca(serie_prod, n_mois_prev)

    fig5 = go.Figure()
    fig5.add_trace(go.Scatter(x=prod_mois["Mois_Label"], y=prod_mois["Production réelle"],
        name="Production réelle", mode="lines+markers",
        line=dict(color=C["violet"], width=2.5), marker=dict(size=5, color=C["violet"]),
        fill="tozeroy", fillcolor="rgba(81,69,158,0.08)",
        hovertemplate="<b>%{x}</b><br>Production réelle : %{y:,.0f} u.<extra></extra>"))
    fig5.add_trace(go.Scatter(x=list(prev_prod["Période"])+list(prev_prod["Période"])[::-1],
        y=list(prev_prod["Haut"])+list(prev_prod["Bas"])[::-1],
        fill="toself", fillcolor="rgba(247,179,49,0.12)",
        line=dict(color="rgba(0,0,0,0)"), showlegend=True, name="Fourchette probable",
        hoverinfo="skip"))
    fig5.add_trace(go.Scatter(x=prev_prod["Période"], y=prev_prod["CA_Prévu"],
        name="Prévision production", mode="lines+markers",
        line=dict(color=C["orange"], width=3, dash="dash"),
        marker=dict(size=8, symbol="diamond", color=C["orange"], line=dict(color="white",width=2)),
        hovertemplate="<b>Prévision %{x}</b><br>%{y:,.0f} unités prévues<extra></extra>"))
    fig5 = excel_style(fig5, 320)
    fig5.update_xaxes(tickangle=-35, tickfont=dict(size=8))
    fig5.update_yaxes(tickformat=",")
    st.plotly_chart(fig5, use_container_width=True)

    prod_prev_moy = prev_prod["CA_Prévu"].mean()
    prod_hist_moy = prod_mois["Production réelle"].mean()
    evol_prod = (prod_prev_moy - prod_hist_moy) / prod_hist_moy * 100 if prod_hist_moy else 0
    c1,c2,c3 = st.columns(3)
    with c1: prev_kpi("⚙️","Production Totale Prévue", f"{prev_prod['CA_Prévu'].sum():,.0f}", f"Sur {n_mois_prev} mois")
    with c2: prev_kpi("📊","Évolution vs Historique", f"{evol_prod:+.1f}%", "Tendance de production")
    with c3: prev_kpi("🏭","Moy. Mensuelle Prévue", f"{prod_prev_moy:,.0f}", "Unités / mois")


# ══════════════════════════════════════════════════════════════



# ══════════════════════════════════════════════════════════════
# PAGE COMPARAISONS & ANALYSES  (insérée avant Alertes)
# ══════════════════════════════════════════════════════════════
elif page == "📊  Comparaisons & Analyses":

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#1E2A4A,#2D3A6B);border-radius:14px;
                padding:22px 28px;margin-bottom:18px;border:1px solid rgba(97,114,243,0.3);">
        <div style="font-size:1.15rem;font-weight:800;color:white;margin-bottom:4px;">
            📊 Comparaisons & Analyses — Tous les segments en face à face
        </div>
        <div style="font-size:0.8rem;color:rgba(255,255,255,0.65);">
            Sélectionnez librement les segments, zones, années et indicateurs à comparer
        </div>
    </div>""", unsafe_allow_html=True)

    comment(
        "Cette page vous permet de comparer librement <b>n'importe quels segments entre eux</b> "
        "sur les indicateurs qui vous intéressent : CA, paiements, nombre de clients, taux de recouvrement. "
        "Vous pouvez comparer <b>2 à 2</b> pour une analyse précise, ou <b>tous ensemble</b> pour une vue globale. "
        "Utilisez les filtres ci-dessous pour personnaliser votre analyse.",
        "📊 Comment utiliser cette page ?"
    )

    # ── PANNEAU DE CONTRÔLE DE COMPARAISON ─────────────────────────
    st.markdown(f'<div class="sec">⚙️  CONFIGUREZ VOTRE COMPARAISON</div>', unsafe_allow_html=True)

    ctrl1, ctrl2, ctrl3 = st.columns([2, 1, 1])
    with ctrl1:
        comp_segments = st.multiselect(
            "🏷️ Segments à comparer (sélectionnez 2 ou plus)",
            ALL_SEGMENTS,
            default=ALL_SEGMENTS[:4],
            help="Choisissez les segments à mettre en face à face. Minimum 2, pas de maximum."
        )
    with ctrl2:
        comp_annees = st.multiselect("🗓 Années", ALL_ANNEES, default=ALL_ANNEES)
    with ctrl3:
        comp_zones = st.multiselect("🗺 Zones", ALL_ZONES, default=ALL_ZONES)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        metrique = st.selectbox("📐 Indicateur principal",
            ["CA Total (FCFA)", "CA Payé (FCFA)", "CA Impayé (FCFA)",
             "Nombre de factures", "Taux de recouvrement (%)",
             "Panier moyen (FCFA)", "Nombre de clients actifs"])
    with col_m2:
        mode_comp = st.radio("🔀 Mode de comparaison",
            ["Tous les segments sélectionnés", "Comparaison 2 à 2"],
            horizontal=True)

    if len(comp_segments) < 2:
        st.warning("⚠️ Sélectionnez au moins 2 segments pour lancer une comparaison.")
        st.stop()

    # Données filtrées pour comparaison
    df_comp = df_fact[
        df_fact["SEGMENT"].isin(comp_segments) &
        df_fact["ANNEE"].isin(comp_annees) &
        df_fact["ZONE"].isin(comp_zones)
    ].copy()

    if df_comp.empty:
        st.warning("Aucune donnée pour ces filtres. Modifiez votre sélection.")
        st.stop()

    df_comp["Période"] = df_comp["DATE"].dt.to_period("M").astype(str)

    # ── Calculer toutes les métriques par segment
    def calc_metrics(df):
        grp = df.groupby("SEGMENT")
        ca = grp["MONTANT TTC"].sum()
        ca_p = df[df["ETAT DE PAIEMENT"]=="Payée"].groupby("SEGMENT")["MONTANT TTC"].sum()
        ca_i = df[df["ETAT DE PAIEMENT"]=="Impayée"].groupby("SEGMENT")["MONTANT TTC"].sum()
        nb_f = grp["MONTANT TTC"].count()
        panier = grp["MONTANT TTC"].mean()
        nb_cl = grp["CLIENT"].nunique()
        taux_r = (ca_p / ca * 100).fillna(0)
        df_out = pd.DataFrame({"CA Total": ca, "CA Payé": ca_p, "CA Impayé": ca_i,
                               "Nb Factures": nb_f, "Panier Moyen": panier,
                               "Nb Clients": nb_cl, "Taux Recouvrement": taux_r}).fillna(0)
        return df_out.reindex(comp_segments).fillna(0)

    metrics_df = calc_metrics(df_comp)

    metrique_col = {
        "CA Total (FCFA)": "CA Total", "CA Payé (FCFA)": "CA Payé",
        "CA Impayé (FCFA)": "CA Impayé", "Nombre de factures": "Nb Factures",
        "Taux de recouvrement (%)": "Taux Recouvrement",
        "Panier moyen (FCFA)": "Panier Moyen", "Nombre de clients actifs": "Nb Clients"
    }
    col_key = metrique_col[metrique]
    is_pct = "%" in metrique

    # ═══════════════════════════════════════
    # MODE 1 : TOUS LES SEGMENTS ENSEMBLE
    # ═══════════════════════════════════════
    if mode_comp == "Tous les segments sélectionnés":

        # ── KPIs de synthèse globale
        section(f"📋  SYNTHÈSE — {len(comp_segments)} segments comparés")
        kpi_cols = st.columns(min(len(comp_segments), 5))
        seg_colors_list = PIE + ["#4472C4","#ED7D31","#A5A5A5","#FFC000"]
        for i, seg in enumerate(comp_segments[:5]):
            val = metrics_df.loc[seg, col_key] if seg in metrics_df.index else 0
            valstr = f"{val:.1f}%" if is_pct else fmt(val,"").strip() if val > 1000 else f"{val:,.0f}"
            with kpi_cols[i]:
                kpi("🏷️", seg[:22], valstr,
                    f"{metrics_df.loc[seg,'Nb Factures']:.0f} factures · {metrics_df.loc[seg,'Nb Clients']:.0f} clients"
                    if seg in metrics_df.index else "",
                    ["bleu","rose","violet","orange","vert","or"][i % 6])

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 1 : Barres horizontales comparaison indicateur principal
        section(f"📊  {metrique.upper()} — Classement de tous les segments")
        comment(f"Chaque barre = un segment. Plus la barre est longue = meilleure performance sur <b>{metrique}</b>. "
                f"Ce graphique permet de voir d'un coup d'œil qui est en tête et qui est en retard.")

        m_sorted = metrics_df[[col_key]].sort_values(col_key, ascending=True)
        clrs = [seg_colors_list[comp_segments.index(s) % len(seg_colors_list)]
                for s in m_sorted.index]
        text_vals = [f"{v:.1f}%" if is_pct else fmt(v,"").strip() for v in m_sorted[col_key]]

        fig = go.Figure(go.Bar(
            x=m_sorted[col_key], y=m_sorted.index,
            orientation="h", marker_color=clrs,
            text=text_vals, textposition="outside", textfont=dict(size=11),
            hovertemplate="<b>%{y}</b><br>" + metrique + " : %{x:,.0f}<extra></extra>"))
        fig = excel_style(fig, max(340, len(comp_segments)*50), False)
        if is_pct:
            fig.update_xaxes(ticksuffix="%")
        else:
            fig.update_xaxes(tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 2 : Part de chaque segment dans le CA total (donut)
        section("🥧  PART DE CHAQUE SEGMENT DANS LE CA TOTAL")
        comment("Ce graphique montre comment le chiffre d'affaires est réparti entre les segments sélectionnés. "
                "Un segment qui occupe une grande part = notre principal moteur. "
                "Si un seul segment domine trop = risque de dépendance.")
        col_d1, col_d2 = st.columns([1.2, 1.8])
        with col_d1:
            ca_seg = metrics_df["CA Total"].sort_values(ascending=False)
            fig2 = go.Figure(go.Pie(
                labels=ca_seg.index, values=ca_seg.values, hole=0.55,
                marker=dict(colors=[seg_colors_list[comp_segments.index(s) % len(seg_colors_list)]
                                    for s in ca_seg.index],
                            line=dict(color="white",width=2)),
                textinfo="label+percent", textfont=dict(size=10),
                hovertemplate="<b>%{label}</b><br>%{value:,.0f} FCFA (%{percent})<extra></extra>"))
            fig2.update_layout(paper_bgcolor="#FFFFFF", height=340,
                margin=dict(t=10,b=30,l=10,r=10),
                legend=dict(orientation="v", x=1.02, y=0.5,
                            font=dict(color=C["texte"],size=9)))
            st.plotly_chart(fig2, use_container_width=True)

        with col_d2:
            # Tableau récap multi-indicateurs
            section("📋  TABLEAU COMPARATIF COMPLET")
            disp = metrics_df.copy()
            for c in ["CA Total","CA Payé","CA Impayé","Panier Moyen"]:
                disp[c] = disp[c].apply(lambda v: f"{v:,.0f} FCFA")
            disp["Nb Factures"] = disp["Nb Factures"].apply(lambda v: f"{int(v)}")
            disp["Nb Clients"]  = disp["Nb Clients"].apply(lambda v: f"{int(v)}")
            disp["Taux Recouvrement"] = disp["Taux Recouvrement"].apply(lambda v: f"{v:.1f}%")
            disp.index.name = "Segment"
            st.dataframe(disp.reset_index(), use_container_width=True, height=340)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 3 : Évolution CA mensuelle — une courbe par segment
        section("📈  ÉVOLUTION MENSUELLE — Une courbe par segment")
        comment("Chaque courbe = un segment. Vous pouvez voir comment chaque segment évolue dans le temps. "
                "Cliquez sur un segment dans la légende pour l'afficher/masquer et comparer plus facilement.")
        ca_mois_seg = df_comp.groupby(["Période","SEGMENT"])["MONTANT TTC"].sum().reset_index()
        fig3 = go.Figure()
        for i, seg in enumerate(comp_segments):
            d = ca_mois_seg[ca_mois_seg["SEGMENT"]==seg]
            c = seg_colors_list[i % len(seg_colors_list)]
            fig3.add_trace(go.Scatter(x=d["Période"], y=d["MONTANT TTC"],
                name=seg, mode="lines+markers",
                line=dict(color=c, width=2.5),
                marker=dict(size=7, color=c, line=dict(color="white",width=1.5)),
                hovertemplate=f"<b>{seg}</b><br>%{{x}}<br>CA : %{{y:,.0f}} FCFA<extra></extra>"))
        fig3 = excel_style(fig3, 380)
        fig3.update_xaxes(tickangle=-35, tickfont=dict(size=8))
        fig3.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 4 : Radar / Spider chart multi-indicateurs
        section("🕸️  PROFIL RADAR — Comparaison multi-indicateurs simultanée")
        comment("Ce graphique en toile d'araignée permet de comparer plusieurs indicateurs en même temps. "
                "Plus la surface couverte est grande = meilleure performance globale du segment. "
                "Idéal pour voir les forces et faiblesses de chaque segment d'un seul regard.")

        # Normaliser pour radar
        radar_cols = ["CA Total","CA Payé","Nb Factures","Panier Moyen","Nb Clients","Taux Recouvrement"]
        radar_labels = ["CA Total","CA Payé","Nb Factures","Panier Moyen","Nb Clients","Recouvrement %"]
        radar_data = metrics_df[radar_cols].copy()
        radar_norm = radar_data.copy()
        for col in radar_cols:
            max_v = radar_data[col].max()
            if max_v > 0:
                radar_norm[col] = radar_data[col] / max_v * 100

        fig4 = go.Figure()
        for i, seg in enumerate(comp_segments[:8]):  # Max 8 sur radar
            if seg not in radar_norm.index: continue
            vals = radar_norm.loc[seg, radar_cols].tolist()
            raw  = radar_data.loc[seg, radar_cols].tolist()
            vals += [vals[0]]  # fermer le polygone
            labels_ext = radar_labels + [radar_labels[0]]
            c = seg_colors_list[i % len(seg_colors_list)]
            fig4.add_trace(go.Scatterpolar(
                r=vals, theta=labels_ext, name=seg,
                fill="toself", fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.15)",
                line=dict(color=c, width=2),
                marker=dict(size=5, color=c),
                hovertemplate=(f"<b>{seg}</b><br>" +
                    "<br>".join([f"{radar_labels[j]}: {raw[j]:,.0f}" for j in range(len(radar_cols))]) +
                    "<extra></extra>"),
            ))
        fig4.update_layout(
            polar=dict(
                bgcolor="#FAFAFA",
                radialaxis=dict(visible=True, range=[0,100], ticksuffix="%",
                                gridcolor="#E2E8F0", tickfont=dict(size=9, color=C["muted"])),
                angularaxis=dict(gridcolor="#E2E8F0", tickfont=dict(size=10, color=C["texte"])),
            ),
            paper_bgcolor="#FFFFFF", height=450,
            legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=C["texte"],size=10),
                        orientation="v", x=1.08, y=0.5),
            margin=dict(t=30,b=30,l=80,r=160),
        )
        st.plotly_chart(fig4, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 5 : Payé vs Impayé par segment — barres côte à côte
        section("💳  PAYÉ VS IMPAYÉ PAR SEGMENT — Qui paie bien, qui paie mal ?")
        comment("Ce graphique sépare pour chaque segment ce qui a été payé (bleu) et ce qui reste impayé (rose). "
                "Les segments avec une grande barre rose ont un risque de trésorerie plus élevé. "
                "Le taux de recouvrement (% payé) est affiché pour chaque segment.")
        fig5 = go.Figure()
        segs_ord = metrics_df.sort_values("CA Total",ascending=False).index.tolist()
        fig5.add_trace(go.Bar(
            x=segs_ord,
            y=[metrics_df.loc[s,"CA Payé"] if s in metrics_df.index else 0 for s in segs_ord],
            name="Payé ✅", marker_color=C["bleu"],
            hovertemplate="<b>%{x}</b><br>CA Payé : %{y:,.0f} FCFA<extra></extra>"))
        fig5.add_trace(go.Bar(
            x=segs_ord,
            y=[metrics_df.loc[s,"CA Impayé"] if s in metrics_df.index else 0 for s in segs_ord],
            name="Impayé ⏳", marker_color=C["rose"],
            hovertemplate="<b>%{x}</b><br>CA Impayé : %{y:,.0f} FCFA<extra></extra>"))
        # Annotations taux recouvrement
        for s in segs_ord:
            if s in metrics_df.index:
                taux_v = metrics_df.loc[s,"Taux Recouvrement"]
                fig5.add_annotation(x=s, y=metrics_df.loc[s,"CA Total"]*1.05,
                    text=f"{taux_v:.0f}%", showarrow=False,
                    font=dict(size=10, color=C["texte"], family="Inter"),
                    bgcolor="rgba(255,255,255,0.8)",
                    bordercolor=C["bleu"] if taux_v>=80 else C["rouge"], borderwidth=1,
                    borderpad=3)
        fig5.update_layout(barmode="stack")
        fig5 = excel_style(fig5, 380)
        fig5.update_xaxes(tickangle=-20, tickfont=dict(size=9))
        fig5.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig5, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 6 : Comparaison annuelle par segment (barres groupées)
        section("📅  ÉVOLUTION ANNUELLE PAR SEGMENT — Qui progresse ?")
        comment("Chaque groupe de barres = une année. À l'intérieur du groupe, chaque couleur = un segment. "
                "Si une barre grandit d'une année à l'autre = le segment est en croissance.")
        ca_ann_seg = df_comp.groupby(["ANNEE","SEGMENT"])["MONTANT TTC"].sum().reset_index()
        fig6 = go.Figure()
        for i, seg in enumerate(comp_segments):
            d = ca_ann_seg[ca_ann_seg["SEGMENT"]==seg]
            c = seg_colors_list[i % len(seg_colors_list)]
            fig6.add_trace(go.Bar(
                x=d["ANNEE"].astype(str), y=d["MONTANT TTC"],
                name=seg, marker_color=c,
                hovertemplate=f"<b>{seg}</b><br>%{{x}}<br>CA : %{{y:,.0f}} FCFA<extra></extra>"))
        fig6.update_layout(barmode="group")
        fig6 = excel_style(fig6, 360)
        fig6.update_xaxes(tickfont=dict(size=11))
        fig6.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig6, use_container_width=True)

    # ═══════════════════════════════════════
    # MODE 2 : COMPARAISON 2 À 2
    # ═══════════════════════════════════════
    else:
        st.markdown("<br>", unsafe_allow_html=True)
        comment("Choisissez <b>deux segments</b> à comparer en détail. Vous verrez leur performance "
                "côte à côte sur tous les indicateurs clés. Changez la paire à tout moment.",
                "🔀 Comparaison 2 à 2")

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            seg_A = st.selectbox("Segment A", comp_segments, index=0)
        with col_s2:
            seg_B = st.selectbox("Segment B", comp_segments,
                                  index=min(1, len(comp_segments)-1))

        if seg_A == seg_B:
            st.warning("Choisissez deux segments différents.")
            st.stop()

        df_A = df_comp[df_comp["SEGMENT"]==seg_A]
        df_B = df_comp[df_comp["SEGMENT"]==seg_B]

        def seg_kpis(df, seg, color):
            ca   = df["MONTANT TTC"].sum()
            paye = df[df["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
            imp  = df[df["ETAT DE PAIEMENT"]=="Impayée"]["MONTANT TTC"].sum()
            taux = paye/ca*100 if ca else 0
            nb_f = len(df); nb_cl = df["CLIENT"].nunique()
            panier = df["MONTANT TTC"].mean() if len(df) else 0
            return {"CA":ca,"Payé":paye,"Impayé":imp,"Taux":taux,
                    "Factures":nb_f,"Clients":nb_cl,"Panier":panier,"color":color,"seg":seg}

        kA = seg_kpis(df_A, seg_A, C["bleu"])
        kB = seg_kpis(df_B, seg_B, C["rose"])

        # ── Header comparatif
        section(f"⚡  {seg_A}  vs  {seg_B}")

        # KPI visuels côte à côte
        for label, keyA, keyB, fmtfn in [
            ("CA Total",   kA["CA"],      kB["CA"],      lambda v: fmt(v,"").strip()),
            ("CA Payé",    kA["Payé"],    kB["Payé"],    lambda v: fmt(v,"").strip()),
            ("CA Impayé",  kA["Impayé"],  kB["Impayé"],  lambda v: fmt(v,"").strip()),
            ("Taux Recouv.",kA["Taux"],   kB["Taux"],    lambda v: f"{v:.1f}%"),
            ("Nb Factures",kA["Factures"],kB["Factures"],lambda v: f"{int(v)}"),
            ("Nb Clients", kA["Clients"], kB["Clients"], lambda v: f"{int(v)}"),
        ]:
            winner_A = keyA >= keyB if label != "CA Impayé" else keyA <= keyB
            col_a_disp, col_mid, col_b_disp = st.columns([2, 1, 2])
            with col_a_disp:
                border = C["vert"] if winner_A else C["muted"]
                st.markdown(f"""
                <div style="background:#FFFFFF;border:2px solid {border};border-radius:10px;
                            padding:12px 16px;text-align:right;">
                    <div style="font-size:0.68rem;color:{C['muted']};text-transform:uppercase;font-weight:700;">{seg_A[:25]}</div>
                    <div style="font-size:1.35rem;font-weight:800;color:{C['bleu']};margin-top:3px;">{fmtfn(keyA)}</div>
                    {"<div style='font-size:0.75rem;color:"+C['vert']+"'>🏆 En tête</div>" if winner_A else ""}
                </div>""", unsafe_allow_html=True)
            with col_mid:
                st.markdown(f"""
                <div style="text-align:center;padding-top:18px;">
                    <div style="font-size:0.68rem;color:{C['muted']};font-weight:700;">{label}</div>
                    <div style="font-size:1rem;color:{C['texte']};font-weight:600;">VS</div>
                </div>""", unsafe_allow_html=True)
            with col_b_disp:
                border = C["vert"] if not winner_A else C["muted"]
                st.markdown(f"""
                <div style="background:#FFFFFF;border:2px solid {border};border-radius:10px;
                            padding:12px 16px;text-align:left;">
                    <div style="font-size:0.68rem;color:{C['muted']};text-transform:uppercase;font-weight:700;">{seg_B[:25]}</div>
                    <div style="font-size:1.35rem;font-weight:800;color:{C['rose']};margin-top:3px;">{fmtfn(keyB)}</div>
                    {"<div style='font-size:0.75rem;color:"+C['vert']+"'>🏆 En tête</div>" if not winner_A else ""}
                </div>""", unsafe_allow_html=True)
            st.markdown("<div style='margin:4px 0;'></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 1 : Courbes évolution CA mensuelle côte à côte
        section("📈  ÉVOLUTION MENSUELLE — Courbes séparées")
        comment(f"Comparez mois par mois les ventes de <b>{seg_A}</b> (bleu) et <b>{seg_B}</b> (rose). "
                "Si une courbe monte plus vite = ce segment est en forte croissance.")
        ca_m_AB = df_comp[df_comp["SEGMENT"].isin([seg_A,seg_B])].groupby(["Période","SEGMENT"])["MONTANT TTC"].sum().reset_index()
        fig_ev = go.Figure()
        for seg, c, dash in [(seg_A, C["bleu"],"solid"), (seg_B, C["rose"],"dash")]:
            d = ca_m_AB[ca_m_AB["SEGMENT"]==seg]
            fig_ev.add_trace(go.Scatter(x=d["Période"], y=d["MONTANT TTC"],
                name=seg, mode="lines+markers",
                line=dict(color=c, width=3, dash=dash),
                marker=dict(size=8, color=c, line=dict(color="white",width=2)),
                fill="tozeroy",
                fillcolor=f"rgba({int(c[1:3],16)},{int(c[3:5],16)},{int(c[5:7],16)},0.08)",
                hovertemplate=f"<b>{seg}</b><br>%{{x}}<br>CA : %{{y:,.0f}} FCFA<extra></extra>"))
        fig_ev = excel_style(fig_ev, 340)
        fig_ev.update_xaxes(tickangle=-35, tickfont=dict(size=8))
        fig_ev.update_yaxes(tickformat=",.0f", ticksuffix=" F")
        st.plotly_chart(fig_ev, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            # ── GRAPHIQUE 2 : Barres groupées annuelles
            section("📅  CA PAR ANNÉE")
            comment(f"Comparez l'évolution annuelle de <b>{seg_A}</b> vs <b>{seg_B}</b>. "
                    "Qui a progressé le plus ? Qui a reculé ?")
            ca_ann_AB = df_comp[df_comp["SEGMENT"].isin([seg_A,seg_B])].groupby(["ANNEE","SEGMENT"])["MONTANT TTC"].sum().reset_index()
            fig_an = go.Figure()
            for seg, c in [(seg_A, C["bleu"]), (seg_B, C["rose"])]:
                d = ca_ann_AB[ca_ann_AB["SEGMENT"]==seg]
                fig_an.add_trace(go.Bar(x=d["ANNEE"].astype(str), y=d["MONTANT TTC"],
                    name=seg, marker_color=c,
                    hovertemplate=f"<b>{seg}</b><br>%{{x}}<br>%{{y:,.0f}} FCFA<extra></extra>"))
            fig_an.update_layout(barmode="group")
            fig_an = excel_style(fig_an, 300)
            fig_an.update_yaxes(tickformat=",.0f", ticksuffix=" F")
            st.plotly_chart(fig_an, use_container_width=True)

        with col_g2:
            # ── GRAPHIQUE 3 : Payé vs Impayé pour A et B
            section("💳  PAYÉ VS IMPAYÉ")
            comment(f"Qui a le meilleur taux de paiement ? "
                    f"{seg_A} : {kA['Taux']:.0f}% · {seg_B} : {kB['Taux']:.0f}%")
            fig_pay = go.Figure()
            for seg, kk, c_p, c_i in [(seg_A, kA, C["bleu"], "#C5CBFB"), (seg_B, kB, C["rose"], "#FFB7D4")]:
                fig_pay.add_trace(go.Bar(x=[seg], y=[kk["Payé"]], name=f"Payé – {seg}",
                    marker_color=c_p, hovertemplate=f"Payé : %{{y:,.0f}} FCFA<extra></extra>"))
                fig_pay.add_trace(go.Bar(x=[seg], y=[kk["Impayé"]], name=f"Impayé – {seg}",
                    marker_color=c_i, hovertemplate=f"Impayé : %{{y:,.0f}} FCFA<extra></extra>"))
            fig_pay.update_layout(barmode="stack")
            fig_pay = excel_style(fig_pay, 300)
            fig_pay.update_yaxes(tickformat=",.0f", ticksuffix=" F")
            st.plotly_chart(fig_pay, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── GRAPHIQUE 4 : Top clients de chaque segment
        col_g3, col_g4 = st.columns(2)
        with col_g3:
            section(f"🏆  TOP CLIENTS — {seg_A[:20]}")
            top_A = df_A.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=True).tail(6).reset_index()
            fig_cA = go.Figure(go.Bar(x=top_A["MONTANT TTC"], y=top_A["CLIENT"],
                orientation="h", marker_color=C["bleu"],
                text=top_A["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
                textposition="outside", textfont=dict(size=10),
                hovertemplate="<b>%{y}</b><br>%{x:,.0f} FCFA<extra></extra>"))
            fig_cA = excel_style(fig_cA, 300, False)
            fig_cA.update_xaxes(tickformat=",.0f")
            st.plotly_chart(fig_cA, use_container_width=True)

        with col_g4:
            section(f"🏆  TOP CLIENTS — {seg_B[:20]}")
            top_B = df_B.groupby("CLIENT")["MONTANT TTC"].sum().sort_values(ascending=True).tail(6).reset_index()
            fig_cB = go.Figure(go.Bar(x=top_B["MONTANT TTC"], y=top_B["CLIENT"],
                orientation="h", marker_color=C["rose"],
                text=top_B["MONTANT TTC"].apply(lambda v: fmt(v,"").strip()),
                textposition="outside", textfont=dict(size=10),
                hovertemplate="<b>%{y}</b><br>%{x:,.0f} FCFA<extra></extra>"))
            fig_cB = excel_style(fig_cB, 300, False)
            fig_cB.update_xaxes(tickformat=",.0f")
            st.plotly_chart(fig_cB, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Synthèse narrative automatique
        section("📝  SYNTHÈSE AUTOMATIQUE DE LA COMPARAISON")
        winner_ca  = seg_A if kA["CA"] >= kB["CA"] else seg_B
        winner_rec = seg_A if kA["Taux"] >= kB["Taux"] else seg_B
        loser_rec  = seg_B if kA["Taux"] >= kB["Taux"] else seg_A
        diff_ca    = abs(kA["CA"]-kB["CA"])
        diff_pct   = diff_ca/max(kA["CA"],kB["CA"])*100 if max(kA["CA"],kB["CA"])>0 else 0
        st.markdown(f"""
        <div class="comment-box">
            <div class="ct">📝 Analyse automatique : {seg_A} vs {seg_B}</div>
            <ul style="margin:6px 0;padding-left:18px;">
                <li><b>CA :</b> <b>{winner_ca}</b> est en tête avec {fmt(max(kA['CA'],kB['CA']),'').strip()} FCFA,
                soit <b>{diff_pct:.0f}%</b> de plus que l'autre segment ({fmt(diff_ca,'').strip()} FCFA d'écart).</li>
                <li><b>Recouvrement :</b> <b>{winner_rec}</b> paie mieux ({max(kA['Taux'],kB['Taux']):.0f}% payé).
                <b>{loser_rec}</b> présente un risque plus élevé ({min(kA['Taux'],kB['Taux']):.0f}% payé).</li>
                <li><b>Activité :</b> {seg_A} a généré <b>{kA['Factures']} factures</b> via <b>{kA['Clients']} clients</b>
                — {seg_B} a généré <b>{kB['Factures']} factures</b> via <b>{kB['Clients']} clients</b>.</li>
                <li><b>Panier moyen :</b> {seg_A} : {fmt(kA['Panier'],'').strip()} FCFA/commande
                · {seg_B} : {fmt(kB['Panier'],'').strip()} FCFA/commande.</li>
            </ul>
            <b>Recommandation :</b> Concentrez les efforts commerciaux sur <b>{winner_ca}</b> pour maximiser le CA,
            et renforcez le suivi des paiements chez <b>{loser_rec}</b> pour sécuriser la trésorerie.
        </div>""", unsafe_allow_html=True)


# PAGE 9 : ALERTES & CONSEILS
# ══════════════════════════════════════════════════════════════
elif page == "⚠️  Alertes & Conseils":
    ca_total  = df_f["MONTANT TTC"].sum()
    ca_paye   = df_f[df_f["ETAT DE PAIEMENT"]=="Payée"]["MONTANT TTC"].sum()
    ca_impaye = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"]["MONTANT TTC"].sum()
    taux_rec  = ca_paye/ca_total*100 if ca_total>0 else 0
    nb_rupt   = (df_inv["Statut"]=="Non disponible").sum()
    nb_faib   = (df_inv["Statut"]=="Stock faible").sum()
    rend_moy  = df_prod["Taux rendement"].mean()

    comment("Cette page regroupe <b>toutes les informations essentielles</b> pour prendre les bonnes décisions. "
            "🔴 Urgent = action immédiate · 🟡 Attention = à surveiller · 🟢 Positif = bonne nouvelle.",
            "⚠️ Comment utiliser cette page ?")

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    def health_card_v2(col, icon, title, valeur, sous, status_txt, color):
        col.markdown(f"""
        <div style="background:#FFFFFF;border-radius:14px;padding:22px;border:1px solid {C['bordure']};
                    border-top:5px solid {color};box-shadow:0 2px 10px rgba(0,0,0,0.07);">
            <div style="font-size:1.8rem;margin-bottom:8px;">{icon}</div>
            <div style="font-size:0.7rem;font-weight:700;color:{C['muted']};text-transform:uppercase;letter-spacing:0.07em;">{title}</div>
            <div style="font-size:1.6rem;font-weight:800;color:{C['texte']};margin:8px 0 3px;">{valeur}</div>
            <div style="font-size:0.78rem;color:{C['muted']};margin-bottom:10px;">{sous}</div>
            <div style="font-size:0.88rem;font-weight:600;color:{color};">{status_txt}</div>
        </div>""", unsafe_allow_html=True)

    t_col = C["vert"] if taux_rec>=80 else (C["orange"] if taux_rec>=65 else C["rouge"])
    s_col = C["vert"] if (nb_rupt+nb_faib)==0 else (C["orange"] if (nb_rupt+nb_faib)<=3 else C["rouge"])
    p_col = C["vert"] if rend_moy>=85 else (C["orange"] if rend_moy>=75 else C["rouge"])

    health_card_v2(c1,"💰","TRÉSORERIE", f"{taux_rec:.1f}%", f"{fmt(ca_impaye)} non encaissés",
        ("🟢 Objectif atteint" if taux_rec>=80 else "🟡 En dessous de l'objectif" if taux_rec>=65 else "🔴 Situation critique"), t_col)
    health_card_v2(c2,"📦","STOCKS", f"{nb_rupt} rupture(s)", f"{nb_faib} stock(s) faible(s)",
        ("🟢 Situation normale" if (nb_rupt+nb_faib)==0 else "🟡 À surveiller" if (nb_rupt+nb_faib)<=3 else "🔴 Situation critique"), s_col)
    health_card_v2(c3,"🏭","PRODUCTION", f"{rend_moy:.1f}%", f"Rebuts : {df_prod['Rebuts'].sum():,.0f} u.",
        ("🟢 Objectif atteint" if rend_moy>=85 else "🟡 En dessous de l'objectif" if rend_moy>=75 else "🔴 Situation critique"), p_col)

    st.markdown("<br>", unsafe_allow_html=True)
    col_al, col_reco = st.columns(2)

    with col_al:
        section("🚨  ALERTES ACTIVES")
        rupt = df_inv[df_inv["Statut"]=="Non disponible"]
        faib = df_inv[df_inv["Statut"]=="Stock faible"]
        if len(rupt)==0 and len(faib)==0:
            st.markdown('<div class="alert-g">✅<div><b>Stocks en bonne santé</b><br>Aucune rupture ni stock critique détecté.</div></div>', unsafe_allow_html=True)
        for _, r in rupt.iterrows():
            st.markdown(f'<div class="alert-r">🔴<div><b>RUPTURE — {r["designation"]}</b><br><small>{r["categorie"]} · 0 unité · ➜ Commander immédiatement</small></div></div>', unsafe_allow_html=True)
        for _, r in faib.head(4).iterrows():
            st.markdown(f'<div class="alert-y">⚠️<div><b>STOCK FAIBLE — {r["designation"]}</b><br><small>{r["categorie"]} · {int(r["Stock final"])} u. restantes (seuil : {int(r["seuil"])} u.) · ➜ Planifier commande</small></div></div>', unsafe_allow_html=True)
        imp = df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"].groupby("CLIENT")["MONTANT TTC"].sum()
        for cl, v in imp.nlargest(4).items():
            st.markdown(f'<div class="alert-r">💸<div><b>IMPAYÉ — {cl}</b><br><small>Créance : <b>{v:,.0f} FCFA</b> · ➜ Appel + lettre de relance</small></div></div>', unsafe_allow_html=True)
        if rend_moy < 80:
            st.markdown(f'<div class="alert-r">⚙️<div><b>RENDEMENT FAIBLE — {rend_moy:.1f}%</b><br><small>Sous 80% · ➜ Audit technique + maintenance machines</small></div></div>', unsafe_allow_html=True)

    with col_reco:
        section("💡  RECOMMANDATIONS STRATÉGIQUES")
        top_seg    = df_f.groupby("SEGMENT")["MONTANT TTC"].sum().idxmax()
        top_zone   = df_f.groupby("ZONE")["MONTANT TTC"].sum().idxmax()
        top_client = df_f.groupby("CLIENT")["MONTANT TTC"].sum().idxmax()
        ca21 = df_fact[df_fact["ANNEE"]==2022]["MONTANT TTC"].sum()
        ca22 = df_fact[df_fact["ANNEE"]==2023]["MONTANT TTC"].sum()
        crois = (ca22-ca21)/ca21*100 if ca21 else 0

        recos = []
        if taux_rec >= 80:
            recos.append(("g","✅ Trésorerie saine", f"Taux de recouvrement {taux_rec:.1f}% ≥ 80%. Maintenez la rigueur de suivi."))
        else:
            recos.append(("r","🔴 Améliorer le recouvrement", f"Taux à {taux_rec:.1f}% (objectif : 80%). Relances à J+30, J+60, J+90."))
        recos.append(("g",f"🏆 Fidéliser {top_client.split()[0]}", f"Meilleur client : <b>{top_client}</b>. Visiter régulièrement, proposer une remise fidélité."))
        recos.append(("g",f"🗺️ Renforcer la {top_zone}", f"Zone la plus rentable. Augmenter la fréquence des visites commerciales."))
        recos.append(("g",f"🏷️ Miser sur {top_seg.split()[0]}", f"Segment <b>{top_seg}</b> = moteur de croissance. Concentrer la prospection sur ce profil."))
        if crois > 0:
            recos.append(("g",f"📈 Croissance +{crois:.1f}%", f"CA en hausse entre 2022 et 2023. Maintenez en fixant des objectifs mensuels ambitieux."))
        else:
            recos.append(("y","📉 Relancer la croissance", f"CA en recul de {crois:.1f}%. Analyser les clients perdus et lancer un plan de reconquête."))
        if nb_faib > 0:
            recos.append(("y","📦 Optimiser les achats stock", f"{nb_faib} références faibles. Mettre en place une commande automatique dès le seuil atteint."))
        recos.append(("g","🔮 Consulter les Prévisions", "Rendez-vous sur la page <b>Prévisions & Anticipations</b> pour anticiper les 6 prochains mois."))

        for typ, titre, msg in recos:
            cls = {"r":"alert-r","y":"alert-y","g":"alert-g"}[typ]
            st.markdown(f'<div class="{cls}"><div><b>{titre}</b><br><small>{msg}</small></div></div>', unsafe_allow_html=True)

    # Export
    st.markdown("<br>", unsafe_allow_html=True)
    section("📥  EXPORT DU RAPPORT COMPLET")
    comment("Téléchargez le rapport Excel complet pour le partager avec votre équipe ou vos auditeurs. "
            "Il contient toutes les données : factures, stock, production et impayés.", "📥 À quoi sert ce rapport ?")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_fact.to_excel(w, sheet_name="Factures", index=False)
        df_inv.to_excel(w, sheet_name="Inventaire", index=False)
        df_ent.to_excel(w, sheet_name="Entrées Stock", index=False)
        df_sor.to_excel(w, sheet_name="Sorties Stock", index=False)
        df_prod.to_excel(w, sheet_name="Production", index=False)
        df_f[df_f["ETAT DE PAIEMENT"]=="Impayée"].to_excel(w, sheet_name="Impayés Prioritaires", index=False)
    buf.seek(0)
    col_dl, _ = st.columns([1, 3])
    with col_dl:
        st.download_button(label="⬇️  Télécharger le Rapport Excel",
            data=buf, file_name=f"MULTIPACK_Rapport_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
