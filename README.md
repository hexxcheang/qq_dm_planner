# DM Optical – document generator

Makes DM Optical's business documents as PDFs, laid out from the company
templates:

- **Quotation** for a customer, priced from the DMO price list
- **Tax Invoice** and **Delivery Order** from a customer's Purchase Order
- **Purchase Order** to the supplier (Hunan Chiopt), priced from the price lists

## What gets filled in

| Invoice                         | Delivery Order               | Comes from               |
|---------------------------------|------------------------------|--------------------------|
| Customer PO No.                 | PO No.                       | customer's PO            |
| Date                            | Date (same as invoice)       | date you pick in the app |
| Invoice No. `IN[YYMMDD-NN]`     | –                            | date + running number    |
| DO No. `DO[YYMMDD-NN]`          | DO No. (same)                | date + running number    |
| Description, Qty, Units, Unit Price, Amount | Qty, Units, Description | customer's PO      |
| Total, Amount Due               | –                            | sum of the amounts              |

`NN` is `01` for the first PO on a date, `02` for the second, and so on. The
app remembers numbers already used in `data/register.json` and suggests the
next one. You can change it before generating.

Everything else (company header, payment terms, salesman, terms & conditions,
Bill To / Ship To) is kept from the template. Bill To / Ship To can be changed
in the app for a different customer.

The item table grows when a PO has more lines than the template has room for,
and the page is set to print on one page.

## Quotation (customer)

In the **Quotation** tab:

1. Enter the customer's company name, address, person in charge and email
   (or pick a customer you quoted before).
2. Type each item's **DM product code** (e.g. `28160140`) or **Chiopt model
   no.** (e.g. `FA0401C`) and the quantity. The app finds the item in the DMO
   price list and prices it at **JM Buying Price (SGD) ÷ 0.65**, rounded to
   cents.
3. The description is the **product code** followed by the item's series
   heading from the price list, e.g. `28160140, 1/1.8" 5MP FA Series`. Items
   the list has no product code for (some accessories and special lenses)
   use the model no. instead, and the app says so.

The quotation gets Quotation No. `Q[YYMMDD-NN]` for today, Valid Till exactly
one month later (same day next month), the items and the Total (SGD). A
typed unit price overrides the list. The quotation template is uploaded once
in the app, like the PO template.

## PO to supplier (Hunan Chiopt)

In the **PO to supplier** tab, type each item's **DM product code** (e.g.
`73620230`) or **Chiopt model no.** (e.g. `ML5010C`) and the quantity:

1. A DM product code is looked up in the DMO price list ("Product Code (DM
   Optical)" → "Model No. (Chiopt)").
2. The model no. is looked up in the Chiopt price list, and its "DM Buying
   Price (USD)" becomes the unit price (rounded to cents).
3. A model no. typed directly skips step 1.

If a model has no DM buying price (some sheets of the Chiopt list have none),
type the price in the table. A typed price always overrides the list.

The PO gets PO# `PO[YYMMDD-NN]` for the PO date (today unless changed), the
delivery date 10 days later, the items and the Total (USD).

The PO and quotation templates and the two price lists hold buying prices and
contacts, so they are **not** in this repository. Upload them once in the app
(the "Price lists & PO template" box); they are stored on your computer in
the `private` folder, which git ignores. Upload a list again when it changes.

## Start the app

1. Install Python 3.10 or newer from https://www.python.org/downloads/ (one
   time). On Windows, tick **"Add python.exe to PATH"** on the installer's
   first screen.
2. Download this folder (GitHub: **Code → Download ZIP**) and unzip it.
3. Start it:
   - **Windows:** double-click `start_app.bat`.
   - **Mac:** right-click `start_app.command` → **Open**. If macOS refuses,
     open Terminal, type `bash ` (with a space), drag the file into the
     window and press Enter.

   The first start installs what the app needs (a few minutes). The app then
   opens in your browser at http://localhost:8501. Keep the black window open
   while you use it; closing it stops the app.

Optional extras:

- Not needed: an Anthropic API key. The built-in reader is free and works
  offline. It reads text PDFs, scanned PDFs and photos (with OCR), Excel and
  CSV. If you ever want Claude to read unusual PO layouts, set
  `ANTHROPIC_API_KEY` and switch on "Read PO with Claude AI" in the sidebar.

Developers can also run `pip install -r requirements.txt` and
`streamlit run app.py` directly.

## Use

In the page that opens:

1. Upload the customer's PO (PDF, scanned PDF, photo, Excel or CSV).
2. Check the PO number and items it read, and fix anything in the table.
3. Pick the date. The Invoice No. and DO No. update to match.
4. Click **Generate Invoice & DO**, check the previews and download the PDFs.

Copies are also saved in `output/<date>/`, e.g.
`output/2026-09-29/IN260929-01_POR13413.pdf`.

`samples/sample_PO.pdf` and `samples/sample_PO.xlsx` are made-up POs you can
use to try it out.

## Updating the templates

Replace the files in `templates/` with new versions, but keep the cells in
the same places (the positions are listed at the top of
`docgen/generator.py`).

## Development

```
pip install -r requirements-dev.txt
python -m pytest
```

- `docgen/po_parser.py` – reads the PO (built-in rules, or Claude)
- `docgen/ocr.py` – free offline OCR for scanned POs (RapidOCR)
- `docgen/generator.py` – fills the Invoice and DO templates
- `docgen/pdf_render.py` – draws a filled-in template as a PDF (no Excel or
  LibreOffice needed; Calibri is drawn with the metric-compatible Carlito font,
  `docgen/fonts/`, SIL Open Font License)
- `docgen/numbering.py` – document numbers and the register of used numbers
- `docgen/pricing.py` – supplier price lookup (product code → model → USD price)
- `docgen/supplier_po.py` – fills the supplier PO template
- `docgen/quotation.py` – fills the quotation template
- `docgen/private_files.py` – the uploaded price lists / PO template in `private/`
- `app.py` – the Streamlit page
