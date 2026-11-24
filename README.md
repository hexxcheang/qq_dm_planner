# DM Optical – document generator

Makes DM Optical's business documents as PDFs, laid out from the company
templates:

- **Quotation** for a customer, priced from the DMO price masterlist
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

The item table grows when a PO has more lines than the template has room for.
When the items don't fit on one page (about 18 on an invoice, 26 on a DO),
the table continues on the next page at full size: the column headings are
repeated, each page is numbered "Page n of N", continuation pages are headed
"Tax Invoice IN… (continued)" / "Delivery Order DO… (continued)", and the
totals and terms follow the last item.

## Quotation (customer)

In the **Quotation** tab:

1. Enter the customer's company name, address, person in charge and email
   (or pick a customer you quoted before).
2. Type each item's **DM product code** (e.g. `28160140`) or **Chiopt model
   no.** (e.g. `FA0401C`) and the quantity. The app finds the item in the
   price masterlist (`DMO_Lens_Price_List_Masterlist_...xlsx`) and uses its
   **Direct Customer Buying Price (SGD) - overseas customer bear freight cost
   from SG**, rounded to cents. That price fills in by itself in the table;
   type a different Unit Price to give a discount (the preview shows the
   difference from the list price), or clear it to get the list price back.
   Items without a list price (e.g. microscope lenses, lens accessories) ask
   for a typed price.
3. The description is the **product code** followed by the item's series
   heading from the price list, e.g. `28160140, 1/1.8" 5MP FA Series`. Items
   the list has no product code for (some accessories and special lenses)
   use the model no. instead, and the app says so.

The quotation gets Quotation No. `Q[YYMMDD-NN]` for today, Valid Till exactly
one month later (same day next month), the items and the Total (SGD). A
typed unit price overrides the list. The quotation template and the price
masterlist are uploaded once in the app, like the PO template; upload a newer
masterlist or template any time and it is used from then on.

**Signature & company stamp:** upload an image of them once (PNG or JPG; a
white background is made transparent) in the quotation tab's upload box, and
every quotation is signed next to "Prepared by:". Upload a new image any time
to replace it.

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

The PO and quotation templates, the price lists and the signature & stamp
hold buying prices, contacts and signatures, so they are **not** in this
repository. Upload them once in the app; upload one again whenever it changes,
and the new one is used from then on.

## Where your files are kept

Uploaded files, running numbers, remembered customers and the generated PDFs
are kept outside the app folder, so downloading a new version of the app
keeps them:

- `Documents/DM Optical App/Saved files` – templates, price lists, signature
- `Documents/DM Optical App/Records` – running numbers, remembered customers
- `Documents/DM Optical App/Generated documents` – the PDFs, one folder per date

(The home folder is used if there is no Documents folder; set `DMO_APP_HOME`
to choose another place.) Files an earlier version kept inside the app folder
(`private`, `data`, `output`) are copied over automatically the first time.

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

Copies are also saved in `Generated documents/<date>/` (see below), e.g.
`Generated documents/2026-09-29/IN260929-01_POR13413.pdf`.

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
- `docgen/pricing.py` – price lookup: supplier USD prices, and quotation prices
  from the masterlist
- `docgen/supplier_po.py` – fills the supplier PO template
- `docgen/quotation.py` – fills the quotation template
- `docgen/private_files.py` – the uploaded templates, price lists and signature
- `docgen/paths.py` – where saved files, records and PDFs are kept
- `app.py` – the Streamlit page
