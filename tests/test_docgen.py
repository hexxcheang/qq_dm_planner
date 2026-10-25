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
