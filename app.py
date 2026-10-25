"""DM Optical document generator.

Tab 1: upload a customer PO -> check the details -> download the Invoice and
Delivery Order.
Tab 2: list product codes / model numbers -> download a PO for the supplier
(Hunan Chiopt), priced from the price lists.
Tab 3: customer details + product codes / model numbers -> download a
quotation, priced from the DMO price list.

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
from docgen.pricing import QUOTE_FACTOR, DmoCatalog, PriceBook, normalize
from docgen.private_files import (CHIOPT_PRICES, DMO_PRICES, PO_TEMPLATE, QUOTATION_TEMPLATE,
                                  Kind, PrivateStore)
from docgen.quotation import (CustomerBook, Quotation, QuoteCustomer, QuoteLine,
                              generate_quotation, quotation_no)
from docgen.supplier_po import DELIVERY_DAYS, POLine, SupplierPO, generate_supplier_po, po_no

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
REGISTER = Register(ROOT / "data" / "register.json")
PO_REGISTER = Register(ROOT / "data" / "supplier_po_register.json")
QUOTE_REGISTER = Register(ROOT / "data" / "quotation_register.json")
CUSTOMERS = CustomerBook(ROOT / "data" / "quote_customers.json")
STORE = PrivateStore()
ITEM_COLUMNS = ["Description", "Qty", "Units", "Unit Price"]
ORDER_COLUMNS = ["Code", "Qty", "Unit Price (USD)"]
QUOTE_COLUMNS = ["Code", "Qty", "Unit Price (SGD)"]
SUPPLIER_FILES = (PO_TEMPLATE, CHIOPT_PRICES, DMO_PRICES)
QUOTATION_FILES = (QUOTATION_TEMPLATE, DMO_PRICES)


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


def private_files_setup(kinds: tuple[Kind, ...], title: str, key: str) -> None:
    missing = STORE.missing(kinds)
    with st.expander(title + ("" if missing else " ✅"), expanded=bool(missing)):
        st.caption("These files hold buying prices and contacts, so they are kept only on this "
                   "computer (in the app's `private` folder), never online. Upload them once; "
                   "upload again whenever one is updated.")
        for kind in kinds:
            described = STORE.describe(kind)
            mark = "✅" if described else "❌"
            st.markdown(f"{mark} **{kind.label}** – {described or f'not uploaded (e.g. {kind.example})'}")
        uploads = st.file_uploader("Upload the Excel files (you can select several at once)",
                                   type=["xlsx"], accept_multiple_files=True, key=key)
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
                    st.error(f"{up.name} doesn't look like a PO/quotation template or a price list.")
            if changed:
                st.rerun()


def order_frame() -> pd.DataFrame:
    return pd.DataFrame({"Code": [""], "Qty": [1.0], "Unit Price (USD)": [float("nan")]},
                        columns=ORDER_COLUMNS)


def supplier_po_page() -> None:
    st.caption("Supplier: **Hunan Chiopt Optical Technology Co., Ltd** · prices in USD")
    private_files_setup(SUPPLIER_FILES, "Price lists & PO template", "private_upload_po")
    if STORE.missing(SUPPLIER_FILES):
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


# --------------------------------------------------------------------------- Quotation


@st.cache_resource
def load_catalog(dmo: str, _stamp: float) -> DmoCatalog:
    # _stamp (file modification time) makes a newly uploaded list load again.
    return DmoCatalog.from_file(Path(dmo))


def customer_form() -> QuoteCustomer:
    known = CUSTOMERS.all()
    names = [c.company for c in known]
    choice = st.selectbox("Customer", ["➕ New customer"] + names,
                          help="Customers you quoted before are remembered here.")
    picked = known[names.index(choice)] if choice in names else QuoteCustomer()
    k = f"q_{choice}"
    col1, col2 = st.columns(2)
    with col1:
        company = st.text_input("Company name", value=picked.company, key=f"{k}_company")
        address1 = st.text_input("Address line 1", value=picked.address1, key=f"{k}_a1")
        address2 = st.text_input("Address line 2", value=picked.address2, key=f"{k}_a2")
    with col2:
        person = st.text_input("Person in charge", value=picked.person, key=f"{k}_person")
        email = st.text_input("Email", value=picked.email, key=f"{k}_email")
    return QuoteCustomer(company.strip(), address1.strip(), address2.strip(), person.strip(),
                         email.strip())


def quotation_page() -> None:
    st.caption(f"Prices: JM Buying Price (SGD) from the DMO price list ÷ {QUOTE_FACTOR}")
    private_files_setup(QUOTATION_FILES, "Price list & quotation template", "private_upload_quote")
    if STORE.missing(QUOTATION_FILES):
        st.info("Upload the files above to start making quotations.")
        return
    dmo = STORE.path(DMO_PRICES)
    catalog = load_catalog(str(dmo), dmo.stat().st_mtime)

    st.subheader("1. Customer (Attention)")
    customer = customer_form()

    st.subheader("2. Items")
    st.caption("Type a **DM product code** (e.g. 28160140) or a **Chiopt model no.** "
               "(e.g. FA0401C) and the quantity. The quotation shows the product code and its "
               "series. Leave Unit Price empty to use the price list; type a price to override it.")
    edited = st.data_editor(
        pd.DataFrame({"Code": [""], "Qty": [1.0], "Unit Price (SGD)": [float("nan")]},
                     columns=QUOTE_COLUMNS),
        key="q_items", num_rows="dynamic", width="stretch",
        column_config={
            "Code": st.column_config.TextColumn("Product code / Model no.", width="medium"),
            "Qty": st.column_config.NumberColumn(min_value=0, step=1, default=1),
            "Unit Price (SGD)": st.column_config.NumberColumn(
                "Unit Price (SGD) – optional", min_value=0, format="%.2f"),
        },
    )

    lines: list[QuoteLine] = []
    preview, problems = [], []
    for row in edited.itertuples(index=False):
        code, qty, override = row[0], row[1], row[2]
        if blank(code) or not normalize(code):
            continue
        qty = 0.0 if blank(qty) else float(qty)
        found = catalog.resolve(code)
        typed = not blank(override)
        price = float(override) if typed else found.unit_price
        if price is None:
            problems.append(f"{found.code}: {found.note or 'enter the price'}")
        elif qty <= 0:
            problems.append(f"{found.code}: enter a quantity")
        else:
            lines.append(QuoteLine(found.description, qty, price))
        preview.append({
            "Code": found.code,
            "Description (on quotation)": found.description,
            "Found by": found.matched_by or "–",
            "JM Buying Price (SGD)": found.item.jm_price_sgd if found.item else None,
            "Unit Price (SGD)": None if price is None else round(price, 2),
            "Price from": "typed" if typed else (f"list ÷ {QUOTE_FACTOR}" if price else "–"),
            "Amount (SGD)": None if price is None else round(qty * round(price, 2), 2),
            "Note": found.note,
        })
    if preview:
        money = st.column_config.NumberColumn(format="%.2f")
        st.dataframe(pd.DataFrame(preview), hide_index=True, width="stretch",
                     column_config={"JM Buying Price (SGD)": money, "Unit Price (SGD)": money,
                                    "Amount (SGD)": money})
        st.markdown(f"**Total (SGD):** {sum(line.amount for line in lines):,.2f}")

    st.subheader("3. Quotation details")
    col1, col2, col3 = st.columns(3)
    with col1:
        q_date = st.date_input("Quotation date", value=date.today(), format="DD/MM/YYYY",
                               key="q_date")
    with col2:
        seq = st.number_input("Running number for this date", min_value=1, max_value=99, step=1,
                              value=QUOTE_REGISTER.next_seq(q_date), key=f"q_seq_{q_date}")
    draft = Quotation(q_date, int(seq), customer, lines)
    with col3:
        st.markdown(f"**Quotation No.:** `{draft.quotation_no}`  \n"
                    f"**Valid till:** {draft.valid_till:%d %b %Y} (one month)")
    just_made = st.session_state.get("q_generated")
    if seq <= QUOTE_REGISTER.last_seq(q_date) and just_made != quotation_no(q_date, seq):
        st.warning(f"{quotation_no(q_date, seq)} has already been used.")

    st.subheader("4. Generate")
    if not customer.company:
        problems.append("Enter the customer's company name.")
    if not lines and not preview:
        problems.append("Add at least one item.")
    for p in problems:
        st.caption(f"⚠️ {p}")
    if st.button("Generate Quotation", type="primary", disabled=bool(problems)):
        with st.spinner("Creating the quotation…"):
            pdf = generate_quotation(draft, STORE.path(QUOTATION_TEMPLATE),
                                     OUTPUT_DIR / q_date.isoformat())
        QUOTE_REGISTER.record(q_date, int(seq), customer.company)
        CUSTOMERS.remember(customer)
        st.session_state["q_generated"] = draft.quotation_no
        st.session_state["q_pdf"] = str(pdf)

    if st.session_state.get("q_pdf"):
        pdf = Path(st.session_state["q_pdf"])
        st.success(f"Created {st.session_state['q_generated']} (saved in {pdf.parent}).")
        show_pdfs([("Quotation", pdf)])


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

quote_tab, invoice_tab, supplier_tab = st.tabs(
    ["💬 Quotation (customer)", "🧾 Invoice & DO (customer)", "📦 PO to supplier (Chiopt)"])
with quote_tab:
    quotation_page()
with invoice_tab:
    invoice_page(use_claude)
with supplier_tab:
    supplier_po_page()
