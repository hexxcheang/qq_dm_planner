"""Business files kept only on this computer (never committed to git).

The supplier PO and the customer quotation need files that hold buying
prices, contacts and the company's signature: the PO template, the quotation
template, the Chiopt price list, the DMO price list and the signature &
stamp image. They are uploaded once in the app and stored in the app's
"Saved files" folder (see ``paths.py``), which updates of the app keep.
"""

from __future__ import annotations

import json
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .paths import SAVED_FILES
from .pricing import DM_PRICE_HEADER, PRODUCT_CODE_HEADER, workbook_mentions

PRIVATE_DIR = SAVED_FILES

PO_TEMPLATE_MARK = re.compile(r"^total\s*\(usd\)$", re.I)
QUOTATION_TEMPLATE_MARK = re.compile(r"^quotation\s*no\.?$", re.I)


@dataclass(frozen=True)
class Kind:
    key: str
    filename: str
    label: str
    example: str


PO_TEMPLATE = Kind("po_template", "po_template.xlsx", "PO template", "DMO_PO_Template.xlsx")
CHIOPT_PRICES = Kind("chiopt_prices", "chiopt_prices.xlsx", "Chiopt price list",
                     "Lens Price_Chiopt to DM.xlsx")
DMO_PRICES = Kind("dmo_prices", "dmo_prices.xlsx", "DMO price list (product codes)",
                  "DMO Lens Price List_JM Buying Price v5.xlsx")
QUOTATION_TEMPLATE = Kind("quotation_template", "quotation_template.xlsx", "Quotation template",
                          "DMO_Quotation_Template.xlsx")
SIGNATURE = Kind("signature_stamp", "signature_stamp.png", "Signature & company stamp (optional)",
                 "a PNG or JPG of the signature and stamp")
KINDS = (PO_TEMPLATE, QUOTATION_TEMPLATE, CHIOPT_PRICES, DMO_PRICES, SIGNATURE)


def classify(path: Path) -> Kind | None:
    """Which of the private files this is, judged by its contents."""
    if _is_image(path):
        return SIGNATURE
    if workbook_mentions(path, DM_PRICE_HEADER):
        return CHIOPT_PRICES
    if workbook_mentions(path, PRODUCT_CODE_HEADER):
        return DMO_PRICES
    if workbook_mentions(path, QUOTATION_TEMPLATE_MARK, max_rows=60):
        return QUOTATION_TEMPLATE
    if workbook_mentions(path, PO_TEMPLATE_MARK, max_rows=60):
        return PO_TEMPLATE
    return None


class PrivateStore:
    def __init__(self, folder: Path = PRIVATE_DIR):
        self.folder = Path(folder)

    def path(self, kind: Kind) -> Path | None:
        p = self.folder / kind.filename
        return p if p.exists() else None

    def _info(self) -> dict:
        try:
            return json.loads((self.folder / "files.json").read_text())
        except (OSError, json.JSONDecodeError):
            return {}

    def describe(self, kind: Kind) -> str | None:
        """'<original name>, saved <date>' for a stored file, or None."""
        if not self.path(kind):
            return None
        info = self._info().get(kind.key, {})
        return f"{info.get('name', kind.filename)}, saved {info.get('saved', '?')}"

    def save(self, source: Path, original_name: str) -> Kind | None:
        """Store ``source`` as whichever kind it is. Returns the kind, or None if unknown."""
        kind = classify(source)
        if kind is None:
            return None
        self.folder.mkdir(parents=True, exist_ok=True)
        if kind == SIGNATURE:
            save_transparent_png(source, self.folder / kind.filename)
        else:
            shutil.copyfile(source, self.folder / kind.filename)
        info = self._info()
        info[kind.key] = {"name": original_name, "saved": f"{datetime.now():%d %b %Y %H:%M}"}
        (self.folder / "files.json").write_text(json.dumps(info, indent=2))
        return kind

    def missing(self, kinds: tuple[Kind, ...] = KINDS) -> list[Kind]:
        return [k for k in kinds if not self.path(k)]


def _is_image(path: Path) -> bool:
    from PIL import Image

    try:
        with Image.open(path) as img:
            img.verify()
        return True
    except Exception:  # not an image PIL can read
        return False


def save_transparent_png(source: Path, target: Path) -> None:
    """Save the image as PNG; a white background (e.g. a scan or JPG) becomes transparent."""
    from PIL import Image

    with Image.open(source) as img:
        rgba = img.convert("RGBA")
    if rgba.getextrema()[3][0] == 255:  # fully opaque: make the white paper see-through
        pixels = rgba.load()
        width, height = rgba.size
        for y in range(height):
            for x in range(width):
                r, g, b, _ = pixels[x, y]
                ink = 255 - min(r, g, b)  # 0 on white paper, high on ink
                if ink < 20:
                    pixels[x, y] = (255, 255, 255, 0)
                else:
                    alpha = min(255, ink * 2)
                    keep = alpha / 255
                    pixels[x, y] = tuple(int(max(0, min(255, (c - (1 - keep) * 255) / keep)))
                                         for c in (r, g, b)) + (alpha,)
    bbox = rgba.getbbox()
    (rgba.crop(bbox) if bbox else rgba).save(target, "PNG")
