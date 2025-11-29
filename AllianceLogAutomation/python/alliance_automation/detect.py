from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import cv2
import numpy as np
import pytesseract
from PIL import Image

from .image_utils import crop, load_image, scale, threshold, to_cv
from .ocr import find_phrase_center, ocr_text


@dataclass
class Detection:
    found: bool
    x: int = 0
    y: int = 0
    score: float = 0.0
    note: str = ""


def _match_template(image_path: Path, template_path: Path, threshold_score: float) -> Optional[Detection]:
    img = cv2.imread(str(image_path))
    tpl = cv2.imread(str(template_path))
    if img is None or tpl is None:
        return None
    res = cv2.matchTemplate(img, tpl, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    if max_val >= threshold_score:
        h, w, _ = tpl.shape
        cx = max_loc[0] + w // 2
        cy = max_loc[1] + h // 2
        return Detection(True, cx, cy, max_val, note="template")
    return None


def find_close_button(image_path: Path, tesseract_path: Path, use_tesseract: bool = True) -> Detection:
    for phrase in ("X", "x", "Close", "Skip"):
        center = find_phrase_center(image_path, phrase, tesseract_path if use_tesseract else None)
        if center:
            return Detection(True, center[0], center[1], note=f"ocr:{phrase}")

    img = load_image(image_path)
    w, h = img.size

    # Top-right crop (30% width/height)
    crop_w = int(w * 0.30)
    crop_h = int(h * 0.30)
    crop_x = w - crop_w
    crop_y = 0
    tr = crop(img, (crop_x, crop_y, crop_w, crop_h))
    scaled_tr = scale(tr, 2.0)

    for level in (None, 0.3, 0.4, 0.5, 0.6, 0.7):
        candidate = scaled_tr if level is None else threshold(scaled_tr, level)
        if use_tesseract:
            pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)
        text = pytesseract.image_to_string(candidate, config="--psm 6")
        if any(sym in text for sym in ["+", "X", "x"]):
            cx = crop_x + crop_w // 2
            cy = crop_y + crop_h // 2
            return Detection(True, cx, cy, note=f"crop-top-right level={level}")

    # Bottom-center crop (60% width, bottom 15% height)
    crop_w = int(w * 0.60)
    crop_h = int(h * 0.15)
    crop_x = (w - crop_w) // 2
    crop_y = h - crop_h
    bc = crop(img, (crop_x, crop_y, crop_w, crop_h))
    scaled_bc = scale(bc, 2.0)

    for level in (None, 0.3, 0.4, 0.5, 0.6, 0.7):
        candidate = scaled_bc if level is None else threshold(scaled_bc, level)
        if use_tesseract:
            pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)
        text = pytesseract.image_to_string(candidate, config="--psm 6")
        if any(sym.lower() in text.lower() for sym in ["close", "x"]):
            cx = crop_x + crop_w // 2
            cy = crop_y + crop_h // 2
            return Detection(True, cx, cy, note=f"crop-bottom-center level={level}")

    return Detection(False)


def alliance_visible(
    image_path: Path,
    tesseract_path: Path,
    use_tesseract: bool = True,
    alliance_template: Optional[Path] = None,
    template_threshold: float = 0.78,
) -> Detection:
    center = find_phrase_center(image_path, "Alliance", tesseract_path if use_tesseract else None)
    if center:
        return Detection(True, center[0], center[1], note="ocr:alliance")

    if alliance_template:
        tpl_hit = _match_template(image_path, alliance_template, template_threshold)
        if tpl_hit:
            tpl_hit.note = "template:alliance"
            return tpl_hit

    return Detection(False)
