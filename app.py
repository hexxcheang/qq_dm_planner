"""DM Optical document generator.

Tab 1: upload a customer PO -> check the details -> download the Invoice and
Delivery Order.
Tab 2: list product codes / model numbers -> download a PO for the supplier
(Hunan Chiopt), priced from the price lists.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from docgen import VERSION
from docgen.generator import INVOICE_TEMPLATE, generate_documents
from docgen.models import Customer, LineItem, PurchaseOrder
from docgen.numbering import Register, do_no, invoice_no
from docgen.po_parser import SUPPORTED_TYPES, ParseError, claude_available, parse_po
from docgen.pricing import PriceBook, normalize
from docgen.private_files import CHIOPT_PRICES, DMO_PRICES, KINDS, PO_TEMPLATE, PrivateStore
from docgen.supplier_po import DELIVERY_DAYS, POLine, SupplierPO, generate_supplier_po, po_no

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
REGISTER = Register(ROOT / "data" / "register.json")
PO_REGISTER = Register(ROOT / "data" / "supplier_po_register.json")
STORE = PrivateStore()
ITEM_COLUMNS = ["Description", "Qty", "Units", "Unit Price"]
ORDER_COLUMNS = ["Code", "Qty", "Unit Price (USD)"]


@st.cache_data
def template_customer() -> Customer:
    ws = load_workbook(INVOICE_TEMPLATE).active
    return Customer(*[str(ws[f"A{r}"].value or "") for r in range(12, 16)])


def pdf_preview(path: Path):
    """First page of a PDF as an image, to show on the page."""
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(str(path))
    try:
        return pdf[0].render(scale=1.5).to_pil()
    finally:
        pdf.close()


def show_pdfs(files: list[tuple[str, Path]]) -> None:
    cols = st.columns(len(files))
    for col, (label, pdf) in zip(cols, files):
        with col:
            st.download_button(f"⬇️ {label} (PDF)", pdf.read_bytes(), file_name=pdf.name,
                               mime="application/pdf", width="stretch")
            st.image(pdf_preview(pdf), width="stretch")


# --------------------------------------------------------------------------- Invoice & DO


def items_frame(po: PurchaseOrder) -> pd.DataFrame:
    rows = [[i.description, i.qty, i.units, i.unit_price] for i in po.items]
    return pd.DataFrame(rows or [["", 1.0, "Pcs", 0.0]], columns=ITEM_COLUMNS)


def blank(value) -> bool:
    return value is None or (not isinstance(value, str) and pd.isna(value))


def frame_items(df: pd.DataFrame) -> tuple[list[LineItem], list[str]]:
    """Line items from the table, and what is missing from incomplete rows."""
    items, problems = [], []
    for row in df.itertuples(index=False):
        desc = "" if blank(row[0]) else str(row[0]).strip()
        if not desc:
            continue
        if blank(row[1]) or float(row[1]) <= 0:
            problems.append(f"{desc}: enter the quantity")
        elif blank(row[3]):
            problems.append(f"{desc}: enter the unit price")
        else:
            units = "" if blank(row[2]) else str(row[2]).strip()
            items.append(LineItem(description=desc, qty=float(row[1]), units=units or "Pcs",
                                  unit_price=float(row[3])))
    return items, problems


def invoice_page(use_claude: bool) -> None:
    st.subheader("1. Upload the customer's Purchase Order")
    upload = st.file_uploader("PO file", type=[t.lstrip(".") for t in SUPPORTED_TYPES])

    if upload is None:
        st.info("Upload a PO (PDF, Excel, CSV or a photo) to begin.")
        return

    data = upload.getvalue()
    file_key = hashlib.sha1(data).hexdigest()
    if (st.session_state.get("file_key") != file_key
            or st.session_state.get("reader_claude") != use_claude):
        with st.spinner("Reading the PO… (scans take a few seconds per page)"):
            try:
                po, reader = parse_po(data, upload.name, use_claude=use_claude)
            except ParseError as exc:
                st.warning(str(exc))
                po, reader = PurchaseOrder(), "manual"
            except Exception as exc:  # network/API problems: let the user carry on by hand
                st.error(f"Could not read the PO automatically: {exc}")
                po, reader = PurchaseOrder(), "manual"
        st.session_state.update(file_key=file_key, reader_claude=use_claude, po=po,
                                reader=reader, generated=None)

    po: PurchaseOrder = st.session_state["po"]
    if st.session_state["reader"] != "manual":
        if po.items:
            st.success(f"Read {len(po.items)} item(s) from the PO "
                       f"({st.session_state['reader']}). Please check them below.")
        else:
            st.warning("No line items found – please type them in below.")

    st.subheader("2. Check the details")
    col1, col2, col3 = st.columns(3)
    with col1:
        po_number = st.text_input("Customer PO No.", value=po.po_number, key=f"po_{file_key}")
    with col2:
        doc_date = st.date_input("Invoice / DO date", value=date.today(), format="DD/MM/YYYY")
    with col3:
        seq = st.number_input("Running number for this date", min_value=1, max_value=99, step=1,
                              value=REGISTER.next_seq(doc_date),
                              key=f"seq_{doc_date}_{file_key}",
                              help="01 for the first PO on this date, 02 for the second, …")
    st.markdown(f"**Invoice No.:** `{invoice_no(doc_date, seq)}` &nbsp;&nbsp; "
                f"**DO No.:** `{do_no(doc_date, seq)}`")
    just_made = st.session_state.get("generated")
    if seq <= REGISTER.last_seq(doc_date) and not (
            just_made and just_made.invoice_no == invoice_no(doc_date, seq)):
        st.warning(f"Number {seq:02d} has already been used for {doc_date:%d %b %Y}.")

    edited = st.data_editor(
        items_frame(po), key=f"items_{file_key}", num_rows="dynamic", width="stretch",
        column_config={
            "Description": st.column_config.TextColumn(width="large"),
            "Qty": st.column_config.NumberColumn(min_value=0, step=1, default=1),
            "Units": st.column_config.TextColumn(default="Pcs"),
            "Unit Price": st.column_config.NumberColumn(min_value=0, format="%.2f"),
        },
    )
    items, problems = frame_items(edited)
    total = round(sum(i.amount for i in items), 2)
    st.markdown(f"**Total / Amount Due:** ${total:,.2f}")

    default_customer = template_customer()
    with st.expander("Bill To / Ship To customer"):
        change_customer = st.checkbox(f"Use a different customer (default: {default_customer.name})")
        if change_customer:
            customer = Customer(
                name=st.text_input("Company", value=default_customer.name),
                address1=st.text_input("Address line 1", value=default_customer.address1),
                address2=st.text_input("Address line 2", value=default_customer.address2),
                attention=st.text_input("Attention", value=default_customer.attention),
            )
        else:
            customer = None

    st.subheader("3. Generate")
    if not po_number.strip():
        problems.append("Enter the Customer PO No.")
    if not items and not problems:
        problems.append("Add at least one line item.")
    for p in problems:
        st.caption(f"⚠️ {p}")

    if st.button("Generate Invoice & DO", type="primary", disabled=bool(problems)):
        final = PurchaseOrder(po_number=po_number.strip(), items=items, customer=customer)
        with st.spinner("Creating documents…"):
            docs = generate_documents(final, doc_date, int(seq), OUTPUT_DIR / doc_date.isoformat())
        REGISTER.record(doc_date, int(seq), final.po_number)
        st.session_state["generated"] = docs

    docs = st.session_state.get("generated")
    if docs:
        st.success(f"Created {docs.invoice_no} and {docs.do_no} "
                   f"(saved in {docs.invoice_pdf.parent}).")
        show_pdfs([("Invoice", docs.invoice_pdf), ("Delivery Order", docs.do_pdf)])


# --------------------------------------------------------------------------- Supplier PO


@st.cache_resource
def load_price_book(chiopt: str, dmo: str, _stamp: tuple) -> PriceBook:
    # _stamp (file modification times) makes a newly uploaded list load again.
    return PriceBook.from_files(Path(chiopt), Path(dmo))


def private_files_setup() -> None:
    missing = STORE.missing()
    with st.expander("Price lists & PO template" + ("" if missing else " ✅"),
                     expanded=bool(missing)):
        st.caption("These files hold buying prices and supplier details, so they are kept only "
                   "on this computer (in the app's `private` folder), never online. Upload them "
                   "once; upload again whenever a list is updated.")
        for kind in KINDS:
            described = STORE.describe(kind)
            mark = "✅" if described else "❌"
            st.markdown(f"{mark} **{kind.label}** – {described or f'not uploaded (e.g. {kind.example})'}")
        uploads = st.file_uploader("Upload the Excel files (you can select all three at once)",
                                   type=["xlsx"], accept_multiple_files=True, key="private_upload")
        if uploads:
            done = st.session_state.setdefault("private_done", set())
            changed = False
            for up in uploads:
                tag = (up.name, hashlib.sha1(up.getvalue()).hexdigest())
                if tag in done:
                    continue
                with tempfile.TemporaryDirectory() as tmp:
                    path = Path(tmp) / "upload.xlsx"
                    path.write_bytes(up.getvalue())
                    kind = STORE.save(path, up.name)
                done.add(tag)
                if kind:
                    st.success(f"Saved {up.name} as the {kind.label}.")
                    changed = True
                else:
                    st.error(f"{up.name} doesn't look like the PO template or either price list.")
            if changed:
                st.rerun()


def order_frame() -> pd.DataFrame:
    return pd.DataFrame({"Code": [""], "Qty": [1.0], "Unit Price (USD)": [float("nan")]},
                        columns=ORDER_COLUMNS)


def supplier_po_page() -> None:
    st.caption("Supplier: **Hunan Chiopt Optical Technology Co., Ltd** · prices in USD")
    private_files_setup()
    if STORE.missing():
        st.info("Upload the files above to start making supplier POs.")
        return

    chiopt, dmo = STORE.path(CHIOPT_PRICES), STORE.path(DMO_PRICES)
    book = load_price_book(str(chiopt), str(dmo), (chiopt.stat().st_mtime, dmo.stat().st_mtime))

    st.subheader("1. What to order")
    st.caption("Type a **DM product code** (e.g. 73620230) or a **Chiopt model no.** "
               "(e.g. ML5010C) and the quantity. You can paste several rows from Excel. "
               "Leave Unit Price empty to use the price list; type a price to override it.")
    edited = st.data_editor(
        order_frame(), key="spo_items", num_rows="dynamic", width="stretch",
        column_config={
            "Code": st.column_config.TextColumn("Product code / Model no.", width="medium"),
            "Qty": st.column_config.NumberColumn(min_value=0, step=1, default=1),
            "Unit Price (USD)": st.column_config.NumberColumn(
                "Unit Price (USD) – optional", min_value=0, format="%.2f"),
        },
    )

    lines: list[POLine] = []
    preview, problems = [], []
    for row in edited.itertuples(index=False):
        code, qty, override = row[0], row[1], row[2]
        if blank(code) or not normalize(code):
            continue
        qty = 0.0 if blank(qty) else float(qty)
        found = book.resolve(code)
        has_override = not blank(override)
        price = float(override) if has_override else found.price_usd
        description = found.model or str(code).strip()
        source = "typed" if has_override else ("price list" if price is not None else "")
        if price is None:
            problems.append(f"{found.code}: {found.note or 'enter the price'}")
        elif qty <= 0:
            problems.append(f"{found.code}: enter a quantity")
        else:
            lines.append(POLine(description, qty, price))
        preview.append({
            "Code": found.code,
            "Model No. (on PO)": description,
            "Found by": found.matched_by or "–",
            "Unit Price (USD)": None if price is None else round(price, 2),
            "Price from": source or "–",
            "Amount (USD)": None if price is None else round(qty * round(price, 2), 2),
            "Note": found.note if price is None else "",
        })
    if preview:
        st.dataframe(pd.DataFrame(preview), hide_index=True, width="stretch",
                     column_config={
                         "Unit Price (USD)": st.column_config.NumberColumn(format="%.2f"),
                         "Amount (USD)": st.column_config.NumberColumn(format="%.2f"),
                     })
        st.markdown(f"**Total (USD):** {sum(line.amount for line in lines):,.2f}")

    st.subheader("2. PO details")
    col1, col2, col3 = st.columns(3)
    with col1:
        po_date = st.date_input("PO date", value=date.today(), format="DD/MM/YYYY", key="spo_date")
    with col2:
        seq = st.number_input("Running number for this date", min_value=1, max_value=99, step=1,
                              value=PO_REGISTER.next_seq(po_date), key=f"spo_seq_{po_date}")
    draft = SupplierPO(po_date, int(seq), lines)
    with col3:
        st.markdown(f"**PO#:** `{draft.po_no}`  \n"
                    f"**Delivery date:** {draft.delivery_date:%d %b %Y} (+{DELIVERY_DAYS} days)")
    just_made = st.session_state.get("spo_generated")
    if seq <= PO_REGISTER.last_seq(po_date) and just_made != po_no(po_date, seq):
        st.warning(f"{po_no(po_date, seq)} has already been used.")

    st.subheader("3. Generate")
    if not lines and not problems:
        problems.append("Add at least one item.")
    for p in problems:
        st.caption(f"⚠️ {p}")
    if st.button("Generate PO", type="primary", disabled=bool(problems)):
        with st.spinner("Creating the PO…"):
            pdf = generate_supplier_po(draft, STORE.path(PO_TEMPLATE),
                                       OUTPUT_DIR / po_date.isoformat())
        PO_REGISTER.record(po_date, int(seq), ", ".join(line.model for line in lines))
        st.session_state["spo_generated"] = draft.po_no
        st.session_state["spo_pdf"] = str(pdf)

    if st.session_state.get("spo_pdf"):
        pdf = Path(st.session_state["spo_pdf"])
        st.success(f"Created {st.session_state['spo_generated']} (saved in {pdf.parent}).")
        show_pdfs([("Purchase Order", pdf)])


# --------------------------------------------------------------------------- page

st.set_page_config(page_title="DM Optical – Documents", page_icon="🧾", layout="wide")
st.title("DM Optical documents")

with st.sidebar:
    st.header("Settings")
    use_claude = st.toggle("Read PO with Claude AI (optional, paid)", value=claude_available(),
                           disabled=not claude_available(),
                           help="Needs ANTHROPIC_API_KEY. Not required: the free built-in reader "
                                "handles PDFs, scans, photos, Excel and CSV. Claude can help with "
                                "unusual PO layouts.")
    if not claude_available():
        st.caption("Using the free built-in reader (works offline, reads scans too).")
    st.divider()
    st.caption(f"Version {VERSION}")

invoice_tab, supplier_tab = st.tabs(["🧾 Invoice & DO (customer)", "📦 PO to supplier (Chiopt)"])
with invoice_tab:
    invoice_page(use_claude)
with supplier_tab:
    supplier_po_page()
