from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

import pytesseract
from PIL import Image


def _set_tesseract(tesseract_path: Path) -> None:
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)


def ocr_text(image_path: Path, tesseract_path: Optional[Path] = None, psm: str = "6") -> str:
    if tesseract_path:
        _set_tesseract(tesseract_path)
    return pytesseract.image_to_string(Image.open(image_path), config=f"--psm {psm}")


def find_phrase_center(image_path: Path, phrase: str, tesseract_path: Optional[Path] = None) -> Optional[Tuple[int, int]]:
    if tesseract_path:
        _set_tesseract(tesseract_path)

    words = [p.strip().lower() for p in phrase.split() if p.strip()]
    if not words:
        return None

    data = pytesseract.image_to_data(Image.open(image_path), output_type=pytesseract.Output.DICT, config="--psm 6")
    n = len(data["text"])
    for i in range(n):
        match = True
        boxes = []
        for j, w in enumerate(words):
            idx = i + j
            if idx >= n:
                match = False
                break
            if data["text"][idx].strip().lower() != w:
                match = False
                break
            boxes.append(
                (
                    data["left"][idx],
                    data["top"][idx],
                    data["width"][idx],
                    data["height"][idx],
                )
            )
        if match and boxes:
            x1 = min(b[0] for b in boxes)
            y1 = min(b[1] for b in boxes)
            x2 = max(b[0] + b[2] for b in boxes)
            y2 = max(b[1] + b[3] for b in boxes)
            cx = (x1 + x2) // 2
            cy = (y1 + y2) // 2
            return cx, cy
    return None
