from datetime import date, datetime

import pytest

import pdfplumber

from docgen.generator import build_delivery_order, build_invoice, generate_documents
from docgen.models import Customer, LineItem, PurchaseOrder
from docgen.numbering import Register, do_no, invoice_no
from docgen import ocr
from docgen.po_parser import ParseError, items_from_rows, parse_with_rules, short_description

from . import sample_po

DAY = date(2026, 9, 29)


def test_document_numbers():
    assert invoice_no(DAY, 1) == "IN260929-01"
    assert do_no(DAY, 12) == "DO260929-12"


def test_register_increments_per_date(tmp_path):
    reg = Register(tmp_path / "register.json")
    assert reg.next_seq(DAY) == 1
    reg.record(DAY, 1, "PO-A")
    reg.record(DAY, 2, "PO-B")
    assert reg.next_seq(DAY) == 3
    assert reg.next_seq(date(2026, 9, 30)) == 1


def _po(n_items: int) -> PurchaseOrder:
    return PurchaseOrder("POR13413", [LineItem(f"Part {i}", i, "Pcs", 10.0 * i)
                                      for i in range(1, n_items + 1)])


@pytest.mark.parametrize("n_items", [1, 2, 3, 8])
def test_invoice_fields(n_items):
    ws = build_invoice(_po(n_items), DAY, 3).active
    assert ws["C5"].value == "POR13413"
    assert ws["H10"].value == datetime(2026, 9, 29)
    assert ws["H11"].value == "IN260929-03"
    assert ws["H12"].value == "DO260929-03"
    for i in range(n_items):
        row = 18 + i
        assert ws[f"A{row}"].value == i + 1
        assert ws[f"B{row}"].value == f"Part {i + 1}"
        assert ws[f"E{row}"].value == i + 1
        assert ws[f"F{row}"].value == "Pcs"
        assert ws[f"G{row}"].value == 10 * (i + 1)
        assert ws[f"H{row}"].value == f"=G{row}*E{row}"
        assert f"B{row}:D{row}" in {r.coord for r in ws.merged_cells.ranges}
    total_row = next(r for r in range(18, ws.max_row + 1) if ws[f"G{r}"].value == "Total")
    assert ws[f"H{total_row}"].value == f"=SUM(H18:H{18 + n_items - 1})"
    assert ws[f"G{total_row + 1}"].value == "Amount Due"
    assert ws[f"H{total_row + 1}"].value == f"=H{total_row}"
    # Terms block still below the totals, still merged.
    terms = [r for r in ws.merged_cells.ranges if r.min_col == 1 and r.max_col == 7]
    assert terms and terms[0].min_row == total_row + 3


@pytest.mark.parametrize("n_items", [1, 5])
def test_delivery_order_fields(n_items):
    ws = build_delivery_order(_po(n_items), DAY, 1).active
    assert ws["G9"].value == datetime(2026, 9, 29)
    assert ws["G10"].value == "DO260929-01"
    assert ws["G11"].value == "POR13413"
    for i in range(n_items):
        row = 16 + i
        assert ws[f"A{row}"].value == i + 1
        assert ws[f"B{row}"].value == i + 1
        assert ws[f"C{row}"].value == "Pcs"
        assert ws[f"D{row}"].value == f"Part {i + 1}"
        assert ws[f"F{row}"].value == "N/A"
    assert ws[f"A{16 + n_items}"].value is None
    received = next(r for r in range(16, ws.max_row + 1)
                    if ws[f"F{r}"].value == "RECEIVED IN GOOD ORDER AND CONDITION")
    assert ws[f"A{received}"].value == "Ship Via"


def test_customer_override():
    po = _po(1)
    po.customer = Customer("ACME PTE LTD", "1 Main St", "Singapore 123456", "Jane")
    inv = build_invoice(po, DAY, 1).active
    do = build_delivery_order(po, DAY, 1).active
    assert [inv[f"A{r}"].value for r in range(12, 16)] == po.customer.lines()
    assert [inv[f"D{r}"].value for r in range(12, 16)] == po.customer.lines()
    assert [do[f"A{r}"].value for r in range(10, 14)] == po.customer.lines()


def _pdf_text(path) -> str:
    with pdfplumber.open(path) as pdf:
        assert len(pdf.pages) == 1
        return pdf.pages[0].extract_text()


def test_generate_pdfs(tmp_path):
    po = PurchaseOrder("POR13495", [LineItem("73620230 (ML5010C)", 2, "Pcs", 85),
                                    LineItem("Optical alignment kit", 1, "Set", 450.5)])
    docs = generate_documents(po, date(2026, 10, 12), 2, tmp_path)
    assert docs.invoice_pdf.name == "IN261012-02_POR13495.pdf"
    assert docs.do_pdf.name == "DO261012-02_POR13495.pdf"
    invoice = _pdf_text(docs.invoice_pdf)
    assert "Description" in invoice and "Desription" not in invoice
    assert "Carriers must also be notified" in invoice
    for text in ("Tax Invoice", "POR13495", "12-Oct-26", "IN261012-02", "DO261012-02",
                 "73620230 (ML5010C)", "170", "450.5", "$620.50"):
        assert text in invoice, text
    do = _pdf_text(docs.do_pdf)
    for text in ("Delivery Order", "POR13495", "12-Oct-26", "DO261012-02", "73620230 (ML5010C)",
                 "Optical alignment kit", "N/A"):
        assert text in do, text
    assert "IN261012" not in do


def _check_sample(po: PurchaseOrder, with_units: bool = True):
    assert po.po_number == sample_po.PO_NUMBER
    assert [(i.description, i.qty, i.unit_price) for i in po.items] == [
        (d, q, p) for d, q, _u, p in sample_po.ITEMS]
    if with_units:
        assert [i.units for i in po.items] == [u for _d, _q, u, _p in sample_po.ITEMS]


def _rules(data: bytes, name: str, expect_ocr: bool = False) -> PurchaseOrder:
    po, used_ocr = parse_with_rules(data, name)
    assert used_ocr == expect_ocr
    return po


def test_rules_parser_pdf_ruled():
    _check_sample(_rules(sample_po.make_pdf(ruled=True), "po.pdf"))


def test_rules_parser_pdf_unruled():
    _check_sample(_rules(sample_po.make_pdf(ruled=False), "po.pdf"))


def test_rules_parser_xlsx():
    _check_sample(_rules(sample_po.make_xlsx(), "po.xlsx"))


def test_rules_parser_csv():
    po = _rules(sample_po.make_csv(), "po.csv")
    _check_sample(po, with_units=False)
    assert all(i.units == "Pcs" for i in po.items)


def test_rules_parser_rejects_unknown_type():
    with pytest.raises(ParseError):
        parse_with_rules(b"hello", "po.docx")


@pytest.mark.skipif(not ocr.available(), reason="rapidocr not installed")
def test_rules_parser_scanned_pdf_uses_ocr():
    _check_sample(_rules(sample_po.make_scanned_pdf(), "po.pdf", expect_ocr=True))


def test_short_description_part_number_and_model():
    text = ("ML5010C\nMACRO LENS, 50MM, 1.1\", C-MOUNT, DM OPTICAL\nP/N:73620230\n"
            "*CUSTOMER: NIXMA TECHNOLOGY (M) SDN. BHD.\nREMARK:\n1) BUYING PRICE OF SGD 85.00")
    assert short_description(text) == "73620230 (ML5010C)"


def test_short_description_drops_remarks():
    text = "Optical alignment kit\nwith carry case\nREMARK:\nShip via FedEx"
    assert short_description(text) == "Optical alignment kit\nwith carry case"
    assert short_description("15370180 (HC2505A)") == "15370180 (HC2505A)"


def test_items_from_rows_multiline_and_table_end():
    rows = [
        ["Purchase Order"],
        ["PO#", "POR13495"],
        ["S No.", "Description", "Quantity", "UOM", "Unit Price", "Amount"],
        ["1", "ML5010C", "2", "Pcs", "85.00", "170.00"],
        ["", "MACRO LENS, 50MM", "", "", "", ""],
        ["", "P/N:73620230", "", "", "", ""],
        ["", "REMARK:", "", "", "", ""],
        ["", "Tax Code Summary", "Tax Rate", "", "Total Net", "Total Tax"],
        ["", "NT:SG", "0.0%", "", "$170.00", "$0.00"],
        ["S No.", "Description", "Quantity", "UOM", "Unit Price", "Amount"],  # page 2 header
        ["", "", "", "", "", ""],
    ]
    items = items_from_rows(rows)
    assert [(i.description, i.qty, i.units, i.unit_price) for i in items] == [
        ("73620230 (ML5010C)", 2, "Pcs", 85.0)]


@pytest.mark.parametrize("n_items", [1, 18, 19, 26, 27, 34, 50, 120])
def test_long_orders_continue_on_next_pages(tmp_path, n_items):
    po = PurchaseOrder("POR13557", [LineItem(f"8023{i:04d} (MV{i:04d}A)", 3, "Pcs", 10 + i)
                                    for i in range(1, n_items + 1)])
    docs = generate_documents(po, date(2026, 10, 2), 1, tmp_path)
    # Up to 18 items fit on one invoice page and 26 on one DO page, with the totals/terms.
    for path, label, ending, one_page in (
            (docs.invoice_pdf, "Tax Invoice IN261002-01", "Amount Due", 18),
            (docs.do_pdf, "Delivery Order DO261002-01", "RECEIVED", 26)):
        with pdfplumber.open(path) as pdf:
            texts = [page.extract_text() for page in pdf.pages]
            item_words = [w for page in pdf.pages
                          for w in page.extract_words(extra_attrs=["size"])
                          if w["text"].startswith("(MV")]
        count = len(texts)
        # Every item exactly once, in order, at full size (11 pt) on every page.
        assert [w["text"] for w in item_words] == [f"(MV{i:04d}A)" for i in range(1, n_items + 1)]
        assert {round(w["size"]) for w in item_words} == {11}
        assert ending in texts[-1] and all(ending not in t for t in texts[:-1])
        if n_items <= one_page:
            assert count == 1 and "Page 1 of" not in texts[0]
            continue
        assert count >= 2
        for k, text in enumerate(texts, start=1):
            assert f"Page {k} of {count}" in text
            if k > 1:
                assert f"{label} (continued)" in text
                assert "S/N" in text and "Description" in text  # headings repeated
        assert "(MV" in texts[-1]  # the totals page also carries at least one item


def _signature_and_stamp(tmp_path):
    """Stand-in for the saved image: black signature left, blue stamp ring right."""
    from PIL import Image

    img = Image.new("RGBA", (600, 300), (255, 255, 255, 0))
    px = img.load()
    for x in range(10, 200):
        px[x, 150 + (x % 20)] = (0, 0, 0, 255)
    for x in range(300, 600):
        for y in range(300):
            if 120 ** 2 <= (x - 450) ** 2 + (y - 150) ** 2 <= 135 ** 2:
                px[x, y] = (50, 50, 150, 255)
    path = tmp_path / "signature_stamp.png"
    img.save(path)
    return path


def test_company_stamp_is_the_blue_ink_only(tmp_path):
    from PIL import Image

    from docgen.stamp import company_stamp

    stamp = company_stamp(_signature_and_stamp(tmp_path))
    assert stamp.size == (271, 271)  # just the ring, no signature
    black_only = tmp_path / "black.png"
    Image.new("RGBA", (50, 50), (0, 0, 0, 255)).save(black_only)
    assert company_stamp(black_only) is None


@pytest.mark.parametrize("n_items", [1, 50])
def test_invoice_stamped_below_amount_due(tmp_path, n_items):
    po = PurchaseOrder("POR13495", [LineItem(f"Part {i}", 1, "Pcs", 10) for i in range(n_items)])
    docs = generate_documents(po, DAY, 1, tmp_path, stamp_from=_signature_and_stamp(tmp_path))
    with pdfplumber.open(docs.invoice_pdf) as pdf:
        last = pdf.pages[-1]
        due = next(w for w in last.extract_words() if w["text"] == "Due")
        amount = next(w for w in last.extract_words()
                      if w["text"] == "Amount" and abs(w["top"] - due["top"]) < 2)
        stamps = [im for im in last.images if im["top"] > due["bottom"]]
        assert len(stamps) == 1
        assert 0 < stamps[0]["top"] - due["bottom"] < 20  # right below "Amount Due"
        assert stamps[0]["x0"] >= amount["x0"]
        assert round(stamps[0]["x1"] - stamps[0]["x0"]) == 95  # 95 pt wide
        assert all(im["top"] < 120 for page in pdf.pages[:-1] for im in page.images)  # logo only
    with pdfplumber.open(docs.do_pdf) as pdf:
        assert sum(len(page.images) for page in pdf.pages) == 1  # the DO is not stamped


def test_invoice_without_stamp_image(tmp_path):
    docs = generate_documents(_po(1), DAY, 1, tmp_path)
    with pdfplumber.open(docs.invoice_pdf) as pdf:
        assert len(pdf.pages[0].images) == 1  # logo only


def test_combine_purchase_orders():
    from docgen.models import combine_purchase_orders

    a = PurchaseOrder("POR13495", [LineItem("73620230 (ML5010C)", 2, "Pcs", 85)])
    b = PurchaseOrder("POR13413", [LineItem("15370180 (HC2505A)", 4, "Pcs", 93),
                                   LineItem("Lens cap", 4, "Pcs", 2.5)])
    unnumbered = PurchaseOrder("", [LineItem("Adapter", 1, "Pcs", 10)])
    combined = combine_purchase_orders([a, b, unnumbered, a])  # the same PO twice: listed once
    assert combined.po_number == "POR13495, POR13413"
    assert [i.description for i in combined.items] == [
        "73620230 (ML5010C)", "15370180 (HC2505A)", "Lens cap", "Adapter", "73620230 (ML5010C)"]
    assert combined.total == 170 + 372 + 10 + 10 + 170
    combined.items[0].qty = 99
    assert a.items[0].qty == 2  # the original POs are left untouched


@pytest.mark.parametrize("n_pos", [2, 6])
def test_invoice_and_do_for_several_pos(tmp_path, n_pos):
    from docgen.models import combine_purchase_orders

    pos = [PurchaseOrder(f"POR1349{i}", [LineItem(f"7362023{i} (ML501{i}C)", 2, "Pcs", 85 + i)])
           for i in range(n_pos)]
    combined = combine_purchase_orders(pos)
    docs = generate_documents(combined, DAY, 1, tmp_path)
    numbers = [po.po_number for po in pos]
    for path in (docs.invoice_pdf, docs.do_pdf):
        with pdfplumber.open(path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            words = page.extract_words()
        for number in numbers:  # every PO number shown, none overlapping another line
            assert number in text.replace(",", " ").split()
        for item in combined.items:
            assert item.description.split()[0] in text
        overlapping = [(v["text"], w["text"]) for i, v in enumerate(words) for w in words[i + 1:]
                       if v["x0"] < w["x1"] - 1 and w["x0"] < v["x1"] - 1
                       and v["top"] < w["bottom"] - 1 and w["top"] < v["bottom"] - 1]
        assert overlapping == []  # no text drawn over other text
    total = sum(item.amount for item in combined.items)
    with pdfplumber.open(docs.invoice_pdf) as pdf:
        assert f"${total:,.2f}" in pdf.pages[-1].extract_text()


def _overlapping_words(page):
    words = page.extract_words()
    return [(v["text"], w["text"]) for i, v in enumerate(words) for w in words[i + 1:]
            if v["x0"] < w["x1"] - 1 and w["x0"] < v["x1"] - 1
            and v["top"] < w["bottom"] - 1 and w["top"] < v["bottom"] - 1]


LONG_DESCRIPTIONS = [  # from quotation Q261125-01, where lines ran into the next item
    "Prism Block (3 BPF IR 15X15X15mm) with 3 wavelengths - Lambda 1: 1630nm, lambda 2: 1345, "
    "lambda 3: 1195nm",
    "Prism Block (2 BPF IR 15X15X15mm) with 2 wavelengths - Lambda 1: 1630nm, lambda 2: 1195nm",
    "Prism Block (2 BPF IR 15X15X15mm) with 2 wavelengths\n- Lambda 1: 1630nm, lambda 2: 1195nm",
    "Slit, glass type, 50umx18mm",
]


def test_long_descriptions_get_tall_enough_rows(tmp_path):
    po = PurchaseOrder("POR1", [LineItem(d, 1, "Pcs", 100) for d in LONG_DESCRIPTIONS])
    docs = generate_documents(po, DAY, 1, tmp_path)
    for path in (docs.invoice_pdf, docs.do_pdf):
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                assert _overlapping_words(page) == [], path.name
