"""Read a customer Purchase Order into a ``PurchaseOrder``.

Two readers:

* ``parse_with_claude`` sends the PO (PDF, image, Excel or CSV) to Claude and
  gets back structured JSON. It copes with any PO layout, including scans.
  Needs ``ANTHROPIC_API_KEY``.
* ``parse_with_rules`` needs no API key. It reads text-based PDFs, Excel and
  CSV files, finds the PO number by its label and the item table by its
  header row (Description / Qty / Unit / Unit Price / Amount).

Either way the result is shown in the app for checking before anything is
generated.
"""

from __future__ import annotations

import base64
import csv
import io
import os
import re
from pathlib import Path

from pydantic import BaseModel, Field

from .models import LineItem, PurchaseOrder

PDF_TYPES = {".pdf"}
IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".webp": "image/webp", ".gif": "image/gif"}
SHEET_TYPES = {".xlsx", ".xlsm"}
CSV_TYPES = {".csv"}
SUPPORTED_TYPES = sorted(PDF_TYPES | set(IMAGE_TYPES) | SHEET_TYPES | CSV_TYPES)

CLAUDE_MODEL = "claude-opus-5-5"


class ParseError(Exception):
    pass


# --------------------------------------------------------------------------- Claude


class _POLine(BaseModel):
    description: str = Field(description="Item description exactly as written on the PO, "
                                         "including part / model numbers")
    quantity: float
    units: str = Field(description="Unit of measure, e.g. Pcs, Set, Box. Use 'Pcs' if not stated")
    unit_price: float = Field(description="Price per unit, without currency symbol")


class _POExtract(BaseModel):
    po_number: str = Field(description="The customer's purchase order number")
    items: list[_POLine]


_PROMPT = (
    "This is a Purchase Order sent by a customer to our company, DM Optical Pte Ltd "
    "(the supplier/vendor). Extract the PO number and every ordered line item. "
    "Copy descriptions as written on the PO. Quantities and prices are plain numbers. "
    "Do not invent items; skip subtotal, GST/tax, delivery and total rows."
)


def claude_available() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN"))


def _content_block(data: bytes, suffix: str) -> dict:
    if suffix in PDF_TYPES:
        return {"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                               "data": base64.standard_b64encode(data).decode()}}
    if suffix in IMAGE_TYPES:
        return {"type": "image", "source": {"type": "base64", "media_type": IMAGE_TYPES[suffix],
                                            "data": base64.standard_b64encode(data).decode()}}
    rows = _rows_from_sheet(data) if suffix in SHEET_TYPES else _rows_from_csv(data)
    text = "\n".join("\t".join(c for c in row) for row in rows)
    return {"type": "text", "text": f"Purchase Order (spreadsheet, tab-separated):\n{text}"}


def parse_with_claude(data: bytes, filename: str) -> PurchaseOrder:
    import anthropic

    suffix = Path(filename).suffix.lower()
    client = anthropic.Anthropic()
    response = client.beta.messages.parse(
        model=CLAUDE_MODEL,
        max_tokens=16000,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        messages=[{"role": "user", "content": [_content_block(data, suffix),
                                               {"type": "text", "text": _PROMPT}]}],
        output_format=_POExtract,
    )
    if response.stop_reason == "refusal" or response.parsed_output is None:
        raise ParseError("Claude could not read this PO. Please enter the details manually.")
    extract = response.parsed_output
    return PurchaseOrder(
        po_number=extract.po_number.strip(),
        items=[LineItem(description=i.description.strip(), qty=i.quantity,
                        units=(i.units or "Pcs").strip() or "Pcs", unit_price=i.unit_price)
               for i in extract.items],
    )


# --------------------------------------------------------------------------- rules

_PO_LABEL = re.compile(
    r"(?:P\.?\s?O\.?|Purchase\s+Order)\s*(?:No\.?|Number|Num\.?|#)?\s*[:#.]?\s*"
    r"([A-Z0-9][A-Z0-9\-/_.]*\d[A-Z0-9\-/_]*)",
    re.IGNORECASE,
)

_HEADER_KEYS = {
    "description": ("description", "desc", "item description", "particulars", "product",
                    "part no", "part number", "item"),
    "qty": ("qty", "quantity", "order qty", "qty ordered"),
    "units": ("uom", "unit", "units", "u/m"),
    "unit_price": ("unit price", "price", "unit cost", "rate", "u/price", "price/unit"),
    "amount": ("amount", "total", "ext price", "extended price", "line total", "net amount"),
}


def _norm(text) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().lower().rstrip(":.")


def _to_number(text) -> float | None:
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    cleaned = re.sub(r"[^\d.\-]", "", str(text).replace(",", ""))
    if cleaned in ("", "-", ".", "-."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _match_header(row: list[str]) -> dict[str, int] | None:
    cells = [_norm(c) for c in row]
    found: dict[str, int] = {}
    for field in ("unit_price", "amount", "qty", "units", "description"):  # most specific first
        for key in _HEADER_KEYS[field]:  # keys in order of preference
            idx = next((i for i, cell in enumerate(cells)
                        if cell and i not in found.values()
                        and (cell == key or cell.startswith(key + " ") or cell.endswith(" " + key))),
                       None)
            if idx is not None:
                found[field] = idx
                break
    if "description" in found and "qty" in found:
        return found
    return None


def items_from_rows(rows: list[list[str]]) -> list[LineItem]:
    """Find the item table header in ``rows`` and read the item lines below it."""
    for h, row in enumerate(rows):
        cols = _match_header(row)
        if not cols:
            continue
        items: list[LineItem] = []
        for line in rows[h + 1:]:
            get = lambda f: line[cols[f]] if f in cols and cols[f] < len(line) else None  # noqa: E731
            desc = str(get("description") or "").strip()
            qty = _to_number(get("qty"))
            if not desc or qty is None or qty <= 0:
                if items and not any(str(c or "").strip() for c in line):
                    continue
                if items and desc and qty is None and not _to_number(get("amount")):
                    items[-1].description += "\n" + desc  # wrapped description line
                    continue
                if items:
                    break  # end of table (subtotal, notes, ...)
                continue
            price = _to_number(get("unit_price"))
            amount = _to_number(get("amount"))
            if price is None and amount is not None:
                price = round(amount / qty, 4)
            units = str(get("units") or "").strip() or "Pcs"
            items.append(LineItem(description=desc, qty=qty, units=units, unit_price=price or 0.0))
        if items:
            return items
    return []


def find_po_number(text: str) -> str:
    for match in _PO_LABEL.finditer(text):
        return match.group(1).strip(".")
    return ""


def _rows_from_sheet(data: bytes) -> list[list[str]]:
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    rows = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else str(v) for v in row])
    return rows


def _rows_from_csv(data: bytes) -> list[list[str]]:
    text = data.decode("utf-8-sig", errors="replace")
    return [row for row in csv.reader(io.StringIO(text))]


def _pdf_tables_and_text(data: bytes) -> tuple[list[list[list[str]]], str]:
    """Candidate row sets (ruled tables first, then text-aligned columns) and the page text."""
    import pdfplumber

    ruled: list[list[str]] = []
    aligned: list[list[str]] = []
    texts: list[str] = []
    text_strategy = {"vertical_strategy": "text", "horizontal_strategy": "text"}
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
            for table in page.extract_tables():
                ruled.extend([["" if c is None else str(c) for c in r] for r in table])
            for table in page.extract_tables(text_strategy):
                aligned.extend([["" if c is None else str(c) for c in r] for r in table])
    return [ruled, aligned], "\n".join(texts)


def parse_with_rules(data: bytes, filename: str) -> PurchaseOrder:
    suffix = Path(filename).suffix.lower()
    if suffix in PDF_TYPES:
        candidates, text = _pdf_tables_and_text(data)
        if not text.strip():
            raise ParseError("This PDF has no text (it is probably a scan). "
                             "Set ANTHROPIC_API_KEY to read scanned POs, or enter the items manually.")
        items = next((found for rows in candidates if (found := items_from_rows(rows))), [])
        return PurchaseOrder(po_number=find_po_number(text), items=items)
    elif suffix in SHEET_TYPES:
        rows = _rows_from_sheet(data)
        text = "\n".join(" ".join(r) for r in rows)
    elif suffix in CSV_TYPES:
        rows = _rows_from_csv(data)
        text = "\n".join(" ".join(r) for r in rows)
    elif suffix in IMAGE_TYPES:
        raise ParseError("Images need the Claude reader. Set ANTHROPIC_API_KEY, "
                         "or enter the items manually.")
    else:
        raise ParseError(f"Unsupported file type {suffix!r}. Use one of: {', '.join(SUPPORTED_TYPES)}")
    return PurchaseOrder(po_number=find_po_number(text), items=items_from_rows(rows))


def parse_po(data: bytes, filename: str, use_claude: bool | None = None) -> tuple[PurchaseOrder, str]:
    """Parse with Claude when available, else with rules. Returns (po, reader_name)."""
    if use_claude is None:
        use_claude = claude_available()
    if use_claude:
        return parse_with_claude(data, filename), "Claude"
    return parse_with_rules(data, filename), "rules"
