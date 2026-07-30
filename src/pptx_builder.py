"""Construit dynamiquement un catalogue PPTX pour un client : couverture + une
ou plusieurs pages par gamme, avec une carte par produit (photo, nom,
description, prix, allergènes).

Contrairement à pptx_generator.py (qui personnalise un deck fixe existant),
ce module génère les slides de zéro avec python-pptx — nécessaire ici car le
nombre de produits par gamme varie selon les données.

Charte reprise de Layla_Apero_Catalogue_FR.pptx : Georgia italique pour les
titres, palette or/crème (#8B6F47 / #C9A876), noir #1A1A1A, gris #6B6B6B.
"""

from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path

import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm, Emu, Pt

from excel_source import PriceResult, compute_final_price
from image_layout import fit_and_center

logger = logging.getLogger("layla_catalog")

SLIDE_WIDTH = Emu(12192000)
SLIDE_HEIGHT = Emu(6858000)

BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x6B, 0x6B, 0x6B)
GOLD = RGBColor(0x8B, 0x6F, 0x47)
GOLD_LIGHT = RGBColor(0xC9, 0xA8, 0x76)
CREAM = RGBColor(0xF7, 0xF3, 0xEC)

FONT_SERIF = "Georgia"
FONT_SANS = "Calibri"

MARGIN_LEFT = Cm(1.78)
MARGIN_RIGHT = Cm(1.78)
MARGIN_TOP_CONTENT = Cm(3.4)
MARGIN_BOTTOM = Cm(1.0)

GRID_COLUMNS = 3
GRID_ROWS = 2
CARDS_PER_SLIDE = GRID_COLUMNS * GRID_ROWS
GUTTER = Cm(0.6)

PHOTO_HEIGHT = Cm(3.6)
NAME_HEIGHT = Cm(0.75)
DESC_HEIGHT = Cm(1.15)
PRICE_HEIGHT = Cm(0.55)
ALLERGENS_HEIGHT = Cm(0.45)

MAX_DESCRIPTION_CHARS = 160


def format_price(value: float) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def _add_textbox(slide, left, top, width, height):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    return box


def _set_run(paragraph, text, font_name, size_pt, color, bold=False, italic=False):
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    return run


def _add_gold_rule(slide, left, top, width_cm=1.8):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left, top, left + Cm(width_cm), top)
    line.line.color.rgb = GOLD
    line.line.width = Pt(1.2)
    return line


def _fit_image_box(image_path: Path, box_w: int, box_h: int) -> tuple[int, int]:
    with Image.open(image_path) as img:
        w_px, h_px = img.size
    aspect = w_px / h_px
    box_aspect = box_w / box_h
    if aspect > box_aspect:
        width = box_w
        height = int(box_w / aspect)
    else:
        height = box_h
        width = int(box_h * aspect)
    return width, height


def _add_photo_or_placeholder(slide, photo_fichier, photos_dir: Path, left, top, width, height, warnings: list[str], product_label: str):
    photo_path = photos_dir / str(photo_fichier) if photo_fichier and not pd.isna(photo_fichier) else None
    if photo_path and photo_path.exists():
        pic_w, pic_h = _fit_image_box(photo_path, width, height)
        pic_left = left + (width - pic_w) // 2
        pic_top = top + (height - pic_h) // 2
        slide.shapes.add_picture(str(photo_path), pic_left, pic_top, width=pic_w, height=pic_h)
        return

    if photo_fichier and not pd.isna(photo_fichier):
        warnings.append(f"Photo introuvable pour '{product_label}' : {photos_dir / str(photo_fichier)}")
    else:
        warnings.append(f"Aucune photo renseignée pour '{product_label}'")

    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = CREAM
    box.line.color.rgb = GOLD
    box.line.width = Pt(0.75)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _set_run(p, "Photo à venir", FONT_SANS, 8, GREY, italic=True)


# Boîte de destination fixe pour le logo partenaire, à droite du logo Layla :
# même sommet et même hauteur que le logo Layla, largeur plafonnée. Le logo
# est mis à l'échelle (sans déformation) pour tenir dans cette boîte puis
# centré dedans, quelles que soient les dimensions/le ratio du fichier source.
PARTNER_LOGO_MAX_WIDTH = Cm(4.6)
PARTNER_LOGO_GAP = Cm(1.0)


def add_cover_slide(prs: Presentation, client_row: pd.Series, logo_dir: Path, layla_logo_path: Path, generation_date: dt.date, warnings: list[str]) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    logo_top = Cm(1.78)
    logo_height = Cm(2.2)
    layla_left = MARGIN_LEFT

    if layla_logo_path.exists():
        with Image.open(layla_logo_path) as img:
            aspect = img.size[0] / img.size[1]
        layla_width = int(logo_height * aspect)
        slide.shapes.add_picture(str(layla_logo_path), layla_left, logo_top, width=layla_width, height=logo_height)
    else:
        layla_width = Cm(4.0)
        warnings.append(f"Logo Layla introuvable : {layla_logo_path}")

    logo_filename = client_row.get("logo_client_fichier")
    if logo_filename and not pd.isna(logo_filename) and str(logo_filename).strip():
        client_logo_path = logo_dir / str(logo_filename).strip()
        if client_logo_path.exists():
            box_left = layla_left + layla_width + PARTNER_LOGO_GAP
            client_left, client_top, width, height = fit_and_center(
                client_logo_path, box_left, logo_top, PARTNER_LOGO_MAX_WIDTH, logo_height, h_align="left"
            )
            slide.shapes.add_picture(str(client_logo_path), client_left, client_top, width=width, height=height)
        else:
            warnings.append(f"Logo client introuvable dans /logo : {logo_filename}")

    title_top = Cm(7.2)
    title_box = _add_textbox(slide, MARGIN_LEFT, title_top, Cm(30), Cm(2.4))
    p = title_box.text_frame.paragraphs[0]
    _set_run(p, str(client_row.get("nom_client", "")), FONT_SERIF, 40, BLACK, italic=True)

    subtitle_box = _add_textbox(slide, MARGIN_LEFT, title_top + Cm(2.5), Cm(30), Cm(1.3))
    p = subtitle_box.text_frame.paragraphs[0]
    _set_run(p, "Sélection & Tarifs", FONT_SERIF, 20, GREY)

    _add_gold_rule(slide, MARGIN_LEFT, title_top + Cm(4.0), width_cm=3.5)

    edition_box = _add_textbox(slide, MARGIN_LEFT, title_top + Cm(4.4), Cm(30), Cm(0.8))
    p = edition_box.text_frame.paragraphs[0]
    _set_run(
        p,
        f"GÉNÉRÉ LE {generation_date.strftime('%d/%m/%Y').upper()}",
        FONT_SANS, 10.5, GOLD, bold=True,
    )

    footer_box = _add_textbox(slide, MARGIN_LEFT, SLIDE_HEIGHT - Cm(1.4), Cm(30), Cm(0.7))
    p = footer_box.text_frame.paragraphs[0]
    _set_run(p, "layla-patisserie.com", FONT_SANS, 9, GREY)


def _add_gamme_header(slide, gamme_name: str, suite: bool) -> None:
    label_box = _add_textbox(slide, MARGIN_LEFT, Cm(0.9), Cm(30), Cm(0.6))
    p = label_box.text_frame.paragraphs[0]
    _set_run(p, "GAMME", FONT_SANS, 10.5, GOLD, bold=True)

    title_text = gamme_name + (" (suite)" if suite else "")
    title_box = _add_textbox(slide, MARGIN_LEFT, Cm(1.5), Cm(30), Cm(1.3))
    p = title_box.text_frame.paragraphs[0]
    _set_run(p, title_text, FONT_SERIF, 24, BLACK, italic=True)

    _add_gold_rule(slide, MARGIN_LEFT, Cm(2.9), width_cm=2.5)


def _add_product_card(slide, product: pd.Series, price_result: PriceResult, photos_dir: Path, left, top, card_w, warnings: list[str]) -> None:
    photo_top = top
    _add_photo_or_placeholder(
        slide, product.get("photo_fichier"), photos_dir, left, photo_top, card_w, PHOTO_HEIGHT,
        warnings, str(product.get("nom_commercial", product.get("code_article"))),
    )

    y = photo_top + PHOTO_HEIGHT + Cm(0.15)
    name_box = _add_textbox(slide, left, y, card_w, NAME_HEIGHT)
    p = name_box.text_frame.paragraphs[0]
    _set_run(p, str(product.get("nom_commercial", "")), FONT_SERIF, 12, BLACK, bold=True)

    y += NAME_HEIGHT
    desc_box = _add_textbox(slide, left, y, card_w, DESC_HEIGHT)
    description = str(product.get("description") or "").strip()
    if len(description) > MAX_DESCRIPTION_CHARS:
        description = description[:MAX_DESCRIPTION_CHARS].rsplit(" ", 1)[0] + "…"
    p = desc_box.text_frame.paragraphs[0]
    _set_run(p, description, FONT_SANS, 8, GREY)

    y += DESC_HEIGHT
    price_box = _add_textbox(slide, left, y, card_w, PRICE_HEIGHT)
    unit_price = None
    poids_net_kg = product.get("poids_net_kg")
    if poids_net_kg and not pd.isna(poids_net_kg):
        unit_price = price_result.price_per_kg * float(poids_net_kg)
    price_text = f"{format_price(price_result.price_per_kg)} / KG"
    if unit_price is not None:
        format_libelle = product.get("format_libelle")
        suffix = f" ({format_libelle})" if format_libelle and not pd.isna(format_libelle) else ""
        price_text += f"   ·   {format_price(unit_price)}{suffix}"
    p = price_box.text_frame.paragraphs[0]
    _set_run(p, price_text, FONT_SANS, 9.5, GOLD, bold=True)

    y += PRICE_HEIGHT
    allergenes = product.get("allergenes")
    if allergenes and not pd.isna(allergenes) and str(allergenes).strip():
        allerg_box = _add_textbox(slide, left, y, card_w, ALLERGENS_HEIGHT)
        p = allerg_box.text_frame.paragraphs[0]
        _set_run(p, f"Allergènes : {allergenes}", FONT_SANS, 7.5, GREY, italic=True)


def add_gamme_slides(
    prs: Presentation,
    gamme_name: str,
    products: pd.DataFrame,
    code_client: str,
    produits_df: pd.DataFrame,
    relations_df: pd.DataFrame,
    on_date: dt.date,
    photos_dir: Path,
    warnings: list[str],
) -> None:
    products = products.sort_values("ordre_affichage", na_position="last")
    rows = list(products.iterrows())

    content_width = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
    card_w = int((content_width - (GRID_COLUMNS - 1) * GUTTER) / GRID_COLUMNS)
    card_stride = card_w + GUTTER
    card_h = PHOTO_HEIGHT + NAME_HEIGHT + DESC_HEIGHT + PRICE_HEIGHT + ALLERGENS_HEIGHT
    row_stride = card_h + GUTTER

    for chunk_index in range(0, len(rows), CARDS_PER_SLIDE):
        chunk = rows[chunk_index: chunk_index + CARDS_PER_SLIDE]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_gamme_header(slide, gamme_name, suite=chunk_index > 0)

        for position, (_, product) in enumerate(chunk):
            col = position % GRID_COLUMNS
            row = position // GRID_COLUMNS
            left = MARGIN_LEFT + col * card_stride
            top = MARGIN_TOP_CONTENT + row * row_stride

            price_result = compute_final_price(
                product["code_article"], code_client, produits_df, relations_df, on_date
            )
            if price_result is None:
                warnings.append(f"Produit '{product['code_article']}' introuvable au moment du calcul du prix — carte ignorée.")
                continue

            _add_product_card(slide, product, price_result, photos_dir, left, top, card_w, warnings)


def build_catalog_for_client(
    client_row: pd.Series,
    produits_df: pd.DataFrame,
    relations_df: pd.DataFrame,
    photos_dir: Path,
    logo_dir: Path,
    layla_logo_path: Path,
    on_date: dt.date,
) -> tuple[Presentation, list[str]]:
    warnings: list[str] = []

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    add_cover_slide(prs, client_row, logo_dir, layla_logo_path, on_date, warnings)

    gammes = (
        produits_df[["gamme", "ordre_affichage"]]
        .dropna(subset=["gamme"])
        .groupby("gamme", as_index=False)["ordre_affichage"]
        .min()
        .sort_values("ordre_affichage", na_position="last")
    )

    for gamme_name in gammes["gamme"]:
        products_in_gamme = produits_df[produits_df["gamme"] == gamme_name]
        add_gamme_slides(
            prs, gamme_name, products_in_gamme, client_row["code_client"],
            produits_df, relations_df, on_date, photos_dir, warnings,
        )

    return prs, warnings
