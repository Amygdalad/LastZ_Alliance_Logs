from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Tuple

import numpy as np
from PIL import Image, ImageOps

from .adb import adb_path, AdbError


def capture_screenshot(serial: str, memu_root: Path, out_path: Path) -> Path:
    cmd = [str(adb_path(memu_root)), "-s", serial, "exec-out", "screencap", "-p"]
    try:
        res = subprocess.run(cmd, capture_output=True, check=False)
    except FileNotFoundError as exc:
        raise AdbError(f"adb not found at {adb_path(memu_root)}") from exc
    if res.returncode != 0 or not res.stdout:
        raise AdbError(f"Screenshot failed: {res.stderr.decode(errors='ignore')}")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(res.stdout)
    return out_path


def load_image(path: Path) -> Image.Image:
    return Image.open(path).convert("RGB")


def save_image(img: Image.Image, path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path)
    return path


def crop(img: Image.Image, box: Tuple[int, int, int, int]) -> Image.Image:
    x, y, w, h = box
    return img.crop((x, y, x + w, y + h))


def scale(img: Image.Image, factor: float) -> Image.Image:
    w, h = img.size
    return img.resize((int(w * factor), int(h * factor)), Image.LANCZOS)


def threshold(img: Image.Image, level: float) -> Image.Image:
    gray = ImageOps.grayscale(img)
    return gray.point(lambda p: 255 if p > level * 255 else 0)


def to_cv(img: Image.Image) -> np.ndarray:
    # Convert PIL to OpenCV BGR
    return np.array(img)[:, :, ::-1].copy()
