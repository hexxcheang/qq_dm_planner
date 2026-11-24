"""Signature and company-stamp pictures placed on the documents.

The saved "signature & company stamp" image holds both: the signature in
black ink and the round company stamp in blue. Quotations use the whole
image; invoices use the stamp alone, taken out of it by its blue ink.
"""

from __future__ import annotations

import io
from pathlib import Path

from openpyxl.drawing.image import Image as SheetImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.worksheet.worksheet import Worksheet

from .pdf_render import EMU_PER_PT, col_width_pt, row_height_pt

# The signature & stamp image is drawn this wide on every document (126 pt = 4.5 cm),
# its size on DM Optical's signed quotation it was cut from.
SIGNATURE_WIDTH_PT = 126.4


def company_stamp(signature_image: Path):
    """The blue company stamp from the signature & stamp image (a PIL image), or None."""
    from PIL import Image

    with Image.open(signature_image) as img:
        rgba = img.convert("RGBA")
    pixels = rgba.load()
    width, height = rgba.size
    found = False
    for y in range(height):
        for x in range(width):
            r, g, b, a = pixels[x, y]
            if a and b - r > 25:  # blue stamp ink
                found = True
            else:
                pixels[x, y] = (255, 255, 255, 0)
    if not found:
        return None
    return rgba.crop(rgba.getbbox())


def place_image(ws: Worksheet, picture, x: float, y: float, width: float) -> None:
    """Anchor ``picture`` (a path or PIL image) with its top-left at (x, y) points, ``width`` wide."""
    col, col_left = 1, 0.0
    while col_left + col_width_pt(ws, col) <= x:
        col_left += col_width_pt(ws, col)
        col += 1
    row, row_top = 1, 0.0
    while row_top + row_height_pt(ws, row) <= y:
        row_top += row_height_pt(ws, row)
        row += 1

    if isinstance(picture, Path):
        img = SheetImage(str(picture))
    else:  # a PIL image made in memory: openpyxl needs it as file data
        data = io.BytesIO()
        picture.save(data, "PNG")
        img = SheetImage(data)
    height = width * img.height / img.width
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col - 1, colOff=int((x - col_left) * EMU_PER_PT),
                           row=row - 1, rowOff=int((y - row_top) * EMU_PER_PT)),
        ext=XDRPositiveSize2D(cx=int(width * EMU_PER_PT), cy=int(height * EMU_PER_PT)))
    ws.add_image(img)


def column_left(ws: Worksheet, column: str) -> float:
    from openpyxl.utils import column_index_from_string

    return sum(col_width_pt(ws, c) for c in range(1, column_index_from_string(column)))


def row_bottom(ws: Worksheet, row: int) -> float:
    return sum(row_height_pt(ws, r) for r in range(1, row + 1))
