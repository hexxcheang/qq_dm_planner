"""Free, offline OCR for scanned POs and photos.

Uses RapidOCR, which installs with pip and ships its models inside the
package, so nothing else needs installing and nothing is sent anywhere.
"""

from __future__ import annotations

import importlib.util
import io
from dataclasses import dataclass
from statistics import median

RENDER_DPI = 200

_engine = None


@dataclass
class Box:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def height(self) -> float:
        return self.y1 - self.y0


def available() -> bool:
    return importlib.util.find_spec("rapidocr") is not None


def _get_engine():
    global _engine
    if _engine is None:
        from rapidocr import RapidOCR

        _engine = RapidOCR()
    return _engine


def ocr_image(image) -> list[Box]:
    """OCR a PIL image into text boxes (pixel coordinates)."""
    result = _get_engine()(image.convert("RGB"))
    if result.boxes is None or result.txts is None:
        return []
    boxes = []
    for points, text in zip(result.boxes, result.txts):
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        boxes.append(Box(str(text).strip(), min(xs), min(ys), max(xs), max(ys)))
    return [b for b in boxes if b.text]


def ocr_pdf(data: bytes) -> list[list[Box]]:
    """OCR every page of a PDF. Returns one list of boxes per page."""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(data)
    try:
        return [ocr_image(page.render(scale=RENDER_DPI / 72).to_pil()) for page in pdf]
    finally:
        pdf.close()


def ocr_image_bytes(data: bytes) -> list[Box]:
    from PIL import Image

    return ocr_image(Image.open(io.BytesIO(data)))


def group_lines(boxes: list[Box]) -> list[list[Box]]:
    """Group boxes into text lines (top to bottom), each sorted left to right.

    A box joins the current line when its vertical centre is within half a
    typical text height of the line's centre, which tolerates slightly skewed
    scans.
    """
    if not boxes:
        return []
    tolerance = median(b.height for b in boxes) / 2
    lines: list[list[Box]] = []
    for box in sorted(boxes, key=lambda b: b.yc):
        if lines:
            line = lines[-1]
            centre = sum(b.yc for b in line) / len(line)
            if abs(box.yc - centre) <= tolerance:
                line.append(box)
                continue
        lines.append([box])
    return [sorted(line, key=lambda b: b.x0) for line in lines]
