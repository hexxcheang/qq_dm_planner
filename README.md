# DM Optical – Invoice & DO generator

Upload a customer's Purchase Order and get a filled-in **Tax Invoice** and
**Delivery Order**, built from the company templates in `templates/`.
(Quotation and supplier PO will come later.)

## What gets filled in

| Invoice                         | Delivery Order               | Comes from               |
|---------------------------------|------------------------------|--------------------------|
| Customer PO No.                 | PO No.                       | customer's PO            |
| Date                            | Date (same as invoice)       | date you pick in the app |
| Invoice No. `IN[YYMMDD-NN]`     | –                            | date + running number    |
| DO No. `DO[YYMMDD-NN]`          | DO No. (same)                | date + running number    |
| Description, Qty, Units, Unit Price, Amount | Qty, Units, Description | customer's PO      |
| Total, Amount Due               | –                            | sum of the amounts (Excel formulas) |

`NN` is `01` for the first PO on a date, `02` for the second, and so on. The
app remembers numbers already used in `data/register.json` and suggests the
next one. You can change it before generating.

Everything else (company header, payment terms, salesman, terms & conditions,
Bill To / Ship To) is kept from the template. Bill To / Ship To can be changed
in the app for a different customer.

The item table grows when a PO has more lines than the template has room for,
and the page is set to print on one page.

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

- Install [LibreOffice](https://www.libreoffice.org) to also get PDF copies
  of the Invoice and DO.
- Set an Anthropic API key so Claude reads the PO. Without it, the built-in
  reader handles text PDFs, Excel and CSV POs that have a normal item table.
  With it, any layout works, including scanned PDFs and photos.

  ```
  # Windows (PowerShell), then restart the app
  setx ANTHROPIC_API_KEY "sk-ant-..."
  # macOS
  export ANTHROPIC_API_KEY="sk-ant-..."
  ```

Developers can also run `pip install -r requirements.txt` and
`streamlit run app.py` directly.

## Use

In the page that opens:

1. Upload the customer's PO (PDF, Excel, CSV, or a photo if Claude is on).
2. Check the PO number and items it read, and fix anything in the table.
3. Pick the date. The Invoice No. and DO No. update to match.
4. Click **Generate Invoice & DO** and download the files.

Copies are also saved in `output/<date>/`, e.g.
`output/2026-09-29/IN260929-01_POR13413.xlsx`.

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

- `docgen/po_parser.py` – reads the PO (Claude or rules)
- `docgen/generator.py` – fills the Invoice and DO templates
- `docgen/numbering.py` – document numbers and the register of used numbers
- `app.py` – the Streamlit page
