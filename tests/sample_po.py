"""Builds fake customer POs (PDF / Excel / CSV) for tests and for trying the app."""

from __future__ import annotations

import io

ITEMS = [
    ("15370180 (HC2505A)", 4, "Pcs", 93.00),
    ("15370181 (HC2506B) Lens holder", 10, "Pcs", 12.50),
    ("Optical alignment kit", 1, "Set", 450.00),
]
PO_NUMBER = "POR13413"


def make_pdf(ruled: bool = True) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet

    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    data = [["No.", "Description", "Qty", "UOM", "Unit Price", "Amount"]]
    for n, (desc, qty, uom, price) in enumerate(ITEMS, start=1):
        data.append([str(n), desc, str(qty), uom, f"{price:,.2f}", f"{qty * price:,.2f}"])
    data.append(["", "", "", "", "Total", f"{sum(q * p for _, q, _, p in ITEMS):,.2f}"])
    table = Table(data)
    if ruled:
        table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    story = [
        Paragraph("JM VISTEC SYSTEM PTE LTD", styles["Title"]),
        Paragraph("PURCHASE ORDER", styles["Heading2"]),
        Paragraph(f"PO No: {PO_NUMBER}", styles["Normal"]),
        Paragraph("PO Date: 25/09/2026", styles["Normal"]),
        Paragraph("Vendor: DM Optical Pte Ltd", styles["Normal"]),
        Spacer(1, 12),
        table,
    ]
    doc.build(story)
    return buf.getvalue()


def make_scanned_pdf() -> bytes:
    """The PDF PO, printed and scanned: pages are pictures only, with no text layer."""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(make_pdf())
    images = [page.render(scale=150 / 72).to_pil().convert("RGB") for page in pdf]
    pdf.close()
    buf = io.BytesIO()
    images[0].save(buf, format="PDF", resolution=150, save_all=True, append_images=images[1:])
    return buf.getvalue()


def make_xlsx() -> bytes:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["PURCHASE ORDER"])
    ws.append(["Purchase Order No.", PO_NUMBER])
    ws.append([])
    ws.append(["S/N", "Item Description", "Quantity", "Unit", "Unit Price (SGD)", "Total (SGD)"])
    for n, (desc, qty, uom, price) in enumerate(ITEMS, start=1):
        ws.append([n, desc, qty, uom, price, qty * price])
    ws.append([None, None, None, None, "Grand Total", sum(q * p for _, q, _, p in ITEMS)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_csv() -> bytes:
    lines = [f"PO #: {PO_NUMBER}", "", "Description,Qty,Price,Amount"]
    for desc, qty, _uom, price in ITEMS:
        lines.append(f"\"{desc}\",{qty},{price},{qty * price}")
    return "\n".join(lines).encode()


if __name__ == "__main__":
    import sys
    from pathlib import Path

    out = Path(sys.argv[1] if len(sys.argv) > 1 else "samples")
    out.mkdir(parents=True, exist_ok=True)
    (out / "sample_PO.pdf").write_bytes(make_pdf())
    (out / "sample_PO.xlsx").write_bytes(make_xlsx())
    print(f"Wrote sample POs to {out}/")
