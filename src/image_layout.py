"""Place an image inside a fixed placeholder box, preserving its aspect ratio
and centering it — so images of any resolution/orientation (e.g. partner
logos supplied by clients) render at a consistent visual size and position
instead of being anchored top-left at whatever size their pixel ratio implies.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image


def fit_and_center(
    image_path: Path,
    box_left: int,
    box_top: int,
    box_width: int,
    box_height: int,
    h_align: str = "center",
) -> tuple[int, int, int, int]:
    """Return (left, top, width, height) in EMU for `image_path` scaled to fit
    within the box without distortion, vertically centered, and horizontally
    either centered ("center") or flush against the box's left edge ("left") —
    "left" keeps the gap to whatever sits immediately left of the box constant
    regardless of the image's aspect ratio."""
    with Image.open(image_path) as img:
        img_w_px, img_h_px = img.size

    scale = min(box_width / img_w_px, box_height / img_h_px)
    width = round(img_w_px * scale)
    height = round(img_h_px * scale)

    left = box_left if h_align == "left" else box_left + (box_width - width) // 2
    top = box_top + (box_height - height) // 2
    return left, top, width, height
