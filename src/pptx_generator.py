"""Génère un catalogue PPTX personnalisé à partir du template Layla_Apero_Catalogue_FR.pptx.

Le deck (9 slides) n'est jamais restructuré : seules deux choses varient par client :
  1. La couverture (slide 1) reçoit le logo du client à côté du logo Layla, si renseigné.
  2. Les prix affichés sur les 3 pages "gamme" (slides 4, 5, 6) sont mis à jour avec le
     prix standard, ou le prix négocié du client s'il existe (feuille Relations).
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

import config
from data_access import (
    get_client_logo_filename,
    get_price,
    load_clients,
    load_produits,
    load_relations,
)
from image_layout import fit_and_center
from versioning import next_available_path

LOGO_GAP_CM = 1.0
LOGO_MARGIN_BEFORE_PHOTO_CM = 0.6
LOGO_MAX_WIDTH_CM = 5.0


def _cm_to_emu(value_cm: float) -> int:
    return int(value_cm * 360000)


def find_shape(slide, shape_id: int):
    for shape in slide.shapes:
        if shape.shape_id == shape_id:
            return shape
    return None


def format_price(value: float) -> str:
    return f"{value:.2f}".replace(".", ",")


def set_shape_text(shape, new_text: str) -> None:
    run = shape.text_frame.paragraphs[0].runs[0]
    run.text = new_text


def add_client_logo(slide, layla_shape, photo_left_emu: int, logo_path: Path):
    """Insert the partner logo into a fixed placeholder box anchored to the
    right of the Layla logo (same top, same height) and left of the cover
    photo, scaling it to fit without distortion and centering it in the box —
    so logos of any resolution/aspect ratio land at the same visual size and
    position (matches the Marriott Tunis placement this box was sized from)."""
    gap = _cm_to_emu(LOGO_GAP_CM)
    margin_before_photo = _cm_to_emu(LOGO_MARGIN_BEFORE_PHOTO_CM)

    box_left = layla_shape.left + layla_shape.width + gap
    box_top = layla_shape.top
    box_width = min(photo_left_emu - margin_before_photo - box_left, _cm_to_emu(LOGO_MAX_WIDTH_CM))
    box_height = layla_shape.height

    left, top, width, height = fit_and_center(logo_path, box_left, box_top, box_width, box_height, h_align="left")
    slide.shapes.add_picture(str(logo_path), left, top, width=Emu(width), height=Emu(height))


def generate_catalog(
    code_client: str | None,
    excel_path: Path = config.EXCEL_PATH,
    template_path: Path = config.TEMPLATE_PATH,
    logo_dir: Path = config.LOGO_DIR,
    price_map_path: Path = config.PRICE_MAP_PATH,
    output_dir: Path = config.OUTPUT_DIR,
) -> tuple[Path, list[str]]:
    warnings: list[str] = []

    produits_df = load_produits(excel_path)
    clients_df = load_clients(excel_path)
    relations_df = load_relations(excel_path)
    price_targets = json.loads(price_map_path.read_text(encoding="utf-8"))["targets"]

    prs = Presentation(template_path)

    if code_client:
        logo_filename = get_client_logo_filename(code_client, clients_df)
        if logo_filename:
            logo_path = logo_dir / logo_filename
            if logo_path.exists():
                cover = prs.slides[config.COVER_SLIDE_INDEX]
                layla_shape = find_shape(cover, shape_id=3)
                photo_shape = find_shape(cover, shape_id=9)
                if layla_shape is not None and photo_shape is not None:
                    add_client_logo(cover, layla_shape, photo_shape.left, logo_path)
                else:
                    warnings.append("Repères de mise en page introuvables sur la couverture : logo client non inséré.")
            else:
                warnings.append(f"Fichier logo introuvable dans /logo : {logo_filename}")

    for target in price_targets:
        slide = prs.slides[target["slide_index"]]
        shape = find_shape(slide, shape_id=target["shape_id"])
        if shape is None:
            warnings.append(f"Zone de prix introuvable sur la slide {target['slide_index'] + 1}.")
            continue

        price, negotiated = get_price(target["code_article"], code_client, produits_df, relations_df)
        if price is None:
            warnings.append(
                f"Pas de prix trouvé pour le code_article '{target['code_article']}' "
                f"({target['family_label']}) — prix du template conservé tel quel."
            )
            continue

        set_shape_text(shape, format_price(price))

    output_dir.mkdir(parents=True, exist_ok=True)
    label = code_client if code_client else "STANDARD"
    date_str = datetime.now().strftime("%Y-%m-%d")
    output_path, _version = next_available_path(output_dir, date_str, label, "pptx")
    prs.save(output_path)

    return output_path, warnings
