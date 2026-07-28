"""Interface Streamlit : génère le catalogue Apéritifs personnalisé pour un client.

Charte visuelle alignée sur Layla_Apero_Catalogue_FR.pptx : Georgia italique,
palette or/crème (#8B6F47 / #C9A876), noir #1A1A1A, gris #6B6B6B.

Lancer avec : streamlit run app.py
"""

import base64
import json
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

import config
from catalog_version_history import (
    construire_nom_fichier,
    enregistrer_version,
    obtenir_historique,
    obtenir_prochain_numero_version,
)
from data_access import load_clients, load_produits
from pptx_generator import generate_catalog
from relations_writer import (
    RelationsWriteError,
    clear_negotiated_price,
    get_negotiated_price,
    upsert_negotiated_price,
)

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

if st.session_state.get("selected_client") != selected_client:
    st.session_state["selected_client"] = selected_client
    st.session_state.pop("price_preview", None)

client_row = clients_df[clients_df["code_client"] == selected_client] if selected_client else None
selected_client_name = client_row.iloc[0]["nom_client"] if client_row is not None else None
is_groupe = (
    client_row is None
    or "type_client" not in clients_df.columns
    or client_row.iloc[0].get("type_client") == "Groupe"
)

st.divider()
st.markdown('<span class="layla-label">Prix négociés</span>', unsafe_allow_html=True)

if selected_client is None:
    st.caption("Sélectionnez un client ci-dessus pour gérer ses prix négociés.")
elif not is_groupe:
    st.info(
        "La négociation se fait uniquement au niveau du groupe/enseigne, jamais hôtel par "
        "hôtel — ce client n'est pas éligible."
    )
else:
    produits_df = load_produits(excel_path)
    price_targets = json.loads(config.PRICE_MAP_PATH.read_text(encoding="utf-8"))["targets"]

    form_values = {}
    for target in price_targets:
        code_article = target["code_article"]
        produit_row = produits_df[produits_df["code_article"] == code_article]
        standard_price = float(produit_row.iloc[0]["prix_standard_kg"]) if not produit_row.empty else None
        current_negotiated = get_negotiated_price(excel_path, selected_client, code_article)

        st.markdown(f"**{target['family_label']}** &nbsp;·&nbsp; *{code_article}*")
        col_standard, col_toggle, col_value = st.columns([1, 1, 2])
        with col_standard:
            st.caption("Prix standard")
            st.write(f"{standard_price:.2f} TND/KG" if standard_price is not None else "—")
        with col_toggle:
            negotiated_checked = st.checkbox(
                "Prix négocié",
                value=current_negotiated is not None,
                key=f"negotiated_toggle_{selected_client}_{code_article}",
            )
        with col_value:
            negotiated_value = st.number_input(
                "Prix négocié (TND/KG)",
                min_value=0.0,
                step=0.10,
                value=current_negotiated if current_negotiated is not None else (standard_price or 0.0),
                key=f"negotiated_value_{selected_client}_{code_article}",
                disabled=not negotiated_checked,
                label_visibility="collapsed",
            )
        form_values[code_article] = {
            "family_label": target["family_label"],
            "checked": negotiated_checked,
            "value": negotiated_value,
            "current": current_negotiated,
        }
        st.write("")

    if st.button("Prévisualiser les changements"):
        preview_rows = []
        for code_article, v in form_values.items():
            current = v["current"]
            if v["checked"]:
                new_value = round(float(v["value"]), 2)
                if current is not None and abs(current - new_value) < 0.005:
                    action = "Aucun changement"
                elif current is None:
                    action = "Créer"
                else:
                    action = "Mettre à jour"
            else:
                new_value = None
                action = "Supprimer (retour au standard)" if current is not None else "Aucun changement"

            preview_rows.append(
                {
                    "Produit": v["family_label"],
                    "code_article": code_article,
                    "Prix actuel": f"{current:.2f}" if current is not None else "Standard",
                    "Nouveau prix": f"{new_value:.2f}" if new_value is not None else "Standard",
                    "Action": action,
                }
            )
        st.session_state["price_preview"] = preview_rows

    preview_rows = st.session_state.get("price_preview")
    if preview_rows:
        st.table(pd.DataFrame(preview_rows).drop(columns=["code_article"]).set_index("Produit"))
        has_changes = any(r["Action"] != "Aucun changement" for r in preview_rows)
        if not has_changes:
            st.caption("Aucun changement à confirmer.")
        elif st.button(f"Confirmer et écrire dans {excel_path.name}", type="primary"):
            try:
                for row in preview_rows:
                    if row["Action"] == "Aucun changement":
                        continue
                    row_code_article = row["code_article"]
                    if row["Action"] == "Supprimer (retour au standard)":
                        clear_negotiated_price(excel_path, selected_client, row_code_article)
                    else:
                        new_price = float(row["Nouveau prix"].replace(",", "."))
                        upsert_negotiated_price(excel_path, selected_client, row_code_article, new_price)
            except RelationsWriteError as exc:
                st.error(str(exc))
            else:
                st.success(
                    f"Prix mis à jour pour {selected_label}. Une sauvegarde du fichier a été "
                    "créée dans backups/."
                )
                st.session_state.pop("price_preview", None)

st.divider()
st.markdown('<span class="layla-label">Langue du catalogue</span>', unsafe_allow_html=True)
selected_language = st.selectbox(
    "Langue",
    options=["Français", "Deutsch (bientôt disponible)", "English (bientôt disponible)"],
    label_visibility="collapsed",
)
if selected_language != "Français":
    st.info(
        "Cette langue sera bientôt disponible (les descriptions produit ne sont pas encore "
        "traduites). Le catalogue généré reste en français pour le moment."
    )

raison = st.text_input(
    "Raison de cette génération (optionnel)",
    placeholder="ex. Nouveau prix négocié, mise à jour du logo...",
)

st.write("")
if st.button("Générer le catalogue", type="primary"):
    n_version = obtenir_prochain_numero_version(
        config.VERSION_HISTORY_PATH, config.CATALOGUE_NAME, client=selected_client_name
    )

    with st.spinner("Génération en cours..."):
        try:
            output_path, warnings = generate_catalog(selected_client, excel_path=excel_path)
        except Exception as exc:
            st.error(f"Échec de la génération : {exc}")
            st.stop()

    # Renomme le fichier produit selon la convention {Catalogue}_{Client}_v{N}_{AAAA-MM-JJ}.pptx
    # et journalise la version, pour garder un historique de chaque génération.
    nom_final = construire_nom_fichier(config.CATALOGUE_NAME, selected_client_name, n_version)
    chemin_final = output_path.with_name(nom_final)
    # .replace() (et non .rename()) : sous Windows, rename() lève une erreur si la
    # cible existe déjà (contrairement à Linux) — replace() écrase proprement.
    output_path.replace(chemin_final)
    output_path = chemin_final

    type_version = "dynamique" if selected_client_name else "statique"
    enregistrer_version(
        config.VERSION_HISTORY_PATH,
        config.CATALOGUE_NAME,
        selected_client_name,
        n_version,
        type_version,
        raison or "Génération depuis l'interface Streamlit",
    )

    st.success(f"Catalogue généré : {output_path.name} (version {n_version})")

    if warnings:
        st.warning("Points à vérifier :\n\n" + "\n".join(f"- {w}" for w in warnings))

    with open(output_path, "rb") as f:
        st.download_button(
            "Télécharger le fichier PPTX",
            data=f.read(),
            file_name=output_path.name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

with st.expander("Historique des versions de ce catalogue"):
    historique = obtenir_historique(config.VERSION_HISTORY_PATH, config.CATALOGUE_NAME, client=selected_client_name)
    if not historique:
        st.caption("Aucune génération enregistrée pour l'instant.")
    else:
        st.table(
            pd.DataFrame(
                [
                    {
                        "Version": e.n_version,
                        "Heure": e.date[11:16],  # ISO "AAAA-MM-JJTHH:MM:SS" -> "HH:MM"
                        "Type": e.type,
                        "Raison": e.raison,
                    }
                    for e in historique
                ]
            ).set_index("Version")
        )

st.divider()
st.markdown(
    f'<span class="layla-footer">Base de données : {excel_path.name} &nbsp;·&nbsp; '
    f'Template : {config.TEMPLATE_PATH.name}</span>',
    unsafe_allow_html=True,
)
