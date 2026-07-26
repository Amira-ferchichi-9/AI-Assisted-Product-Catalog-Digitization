"""Interface Streamlit : génère le catalogue Apéritifs personnalisé pour un client.

Charte visuelle alignée sur Layla_Apero_Catalogue_FR.pptx : Georgia italique,
palette or/crème (#8B6F47 / #C9A876), noir #1A1A1A, gris #6B6B6B.

Lancer avec : streamlit run app.py
"""

import base64
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

import config
from data_access import load_clients
from pptx_generator import generate_catalog

LOGO_PATH = Path(__file__).resolve().parent / "src" / "assets" / "layla_logo.png"

GOLD = "#8B6F47"
GOLD_LIGHT = "#C9A876"
BLACK = "#1A1A1A"
GREY = "#6B6B6B"
CREAM = "#F7F3EC"

st.set_page_config(page_title="Générateur de catalogue Layla", page_icon="🫒", layout="centered")


def _logo_b64() -> str:
    return base64.b64encode(LOGO_PATH.read_bytes()).decode()


st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}

    h1, h2, h3, .layla-title {{
        font-family: Georgia, 'Times New Roman', serif !important;
    }}

    .layla-header {{
        display: flex;
        align-items: center;
        gap: 1.4rem;
        margin-bottom: 0.4rem;
    }}
    .layla-header img {{ height: 46px; }}
    .layla-title {{
        font-style: italic;
        font-weight: 400;
        font-size: 2.1rem;
        color: {BLACK};
        margin: 0;
        line-height: 1.1;
    }}
    .layla-subtitle {{
        font-family: Calibri, 'Segoe UI', sans-serif;
        font-size: 0.8rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: {GOLD};
        font-weight: 600;
        margin: 0.3rem 0 0.9rem 0;
    }}
    .layla-rule {{
        border: none;
        border-top: 1.5px solid {GOLD};
        width: 64px;
        margin: 0 0 1.6rem 0;
    }}
    .layla-label {{
        font-family: Calibri, 'Segoe UI', sans-serif;
        font-size: 0.75rem;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: {GOLD};
        font-weight: 600;
        margin-bottom: -0.6rem;
        margin-top: 1.2rem;
        display: block;
    }}
    .layla-footer {{
        font-family: Calibri, 'Segoe UI', sans-serif;
        font-size: 0.75rem;
        color: {GREY};
    }}

    div[data-testid="stWidgetLabel"] p {{
        font-family: Calibri, 'Segoe UI', sans-serif !important;
        font-size: 0.75rem !important;
        letter-spacing: 0.1em !important;
        text-transform: uppercase !important;
        color: {GOLD} !important;
        font-weight: 600 !important;
    }}

    button[kind="primary"] {{
        background-color: {BLACK} !important;
        color: {GOLD_LIGHT} !important;
        border: none !important;
        border-radius: 2px !important;
        font-family: Calibri, 'Segoe UI', sans-serif !important;
        letter-spacing: 0.08em !important;
        text-transform: uppercase !important;
        font-size: 0.85rem !important;
        padding: 0.6rem 1.6rem !important;
    }}
    button[kind="primary"]:hover {{
        background-color: {GOLD} !important;
        color: #FFFFFF !important;
    }}

    div[data-testid="stAlert"] {{
        font-family: Calibri, 'Segoe UI', sans-serif !important;
        border-radius: 2px !important;
    }}

    hr {{ border-top: 1px solid #E7DFD3 !important; }}
    </style>

    <div class="layla-header">
        <img src="data:image/png;base64,{_logo_b64()}" />
        <div>
            <p class="layla-title">Générateur de catalogue</p>
        </div>
    </div>
    <p class="layla-subtitle">Apéritifs&nbsp;&nbsp;/&nbsp;&nbsp;Personnalisation client</p>
    <hr class="layla-rule" />
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<span class="layla-footer">Personnalise le deck Layla_Apero_Catalogue_FR.pptx : '
    'logo de couverture et prix, selon le client sélectionné.</span>',
    unsafe_allow_html=True,
)

st.markdown('<span class="layla-label">Source de données</span>', unsafe_allow_html=True)
use_sample = st.toggle(
    "Utiliser les données d'exemple (data/sample_produits.xlsx)",
    value=not config.EXCEL_PATH.exists(),
    help="À décocher une fois que votre fichier Excel réel contient les code_article "
         "attendus (voir config/gamme_price_map.json).",
)
excel_path = config.SAMPLE_EXCEL_PATH if use_sample else config.EXCEL_PATH

if not excel_path.exists():
    st.error(f"Fichier introuvable : {excel_path}")
    st.stop()

clients_df = load_clients(excel_path)

client_options = {"— Standard (sans logo client) —": None}
for _, row in clients_df.iterrows():
    label = f"{row['nom_client']} ({row['code_client']})"
    client_options[label] = row["code_client"]

st.markdown('<span class="layla-label">Client</span>', unsafe_allow_html=True)
selected_label = st.selectbox("Client", options=list(client_options.keys()), label_visibility="collapsed")
selected_client = client_options[selected_label]

st.write("")
if st.button("Générer le catalogue", type="primary"):
    with st.spinner("Génération en cours..."):
        try:
            output_path, warnings = generate_catalog(selected_client, excel_path=excel_path)
        except Exception as exc:
            st.error(f"Échec de la génération : {exc}")
            st.stop()

    st.success(f"Catalogue généré : {output_path.name}")

    if warnings:
        st.warning("Points à vérifier :\n\n" + "\n".join(f"- {w}" for w in warnings))

    with open(output_path, "rb") as f:
        st.download_button(
            "Télécharger le fichier PPTX",
            data=f.read(),
            file_name=output_path.name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

st.divider()
st.markdown(
    f'<span class="layla-footer">Base de données : {excel_path.name} &nbsp;·&nbsp; '
    f'Template : {config.TEMPLATE_PATH.name}</span>',
    unsafe_allow_html=True,
)
