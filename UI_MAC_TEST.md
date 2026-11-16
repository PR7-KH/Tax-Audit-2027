# Target-folder batch conversion update

## Install

Download downloads/Tax-Audit-Mac-Update.zip from GitHub. Close the app, back up
your working folder, unzip, and copy its five files into the working app folder
together. Keep your existing templates, company folders and other Python files.
Run `bash Install_Update_Mac.command` from Terminal in the app folder if the PDF
dependencies are not installed, then launch normally. Python 3.10+ is required.

## Menu order

Audit Dashboard → TOIs Overview → MTR Overview → Consolidated data → Tax Paid
Summary. MTR tabs and exports retain their existing calculations.

## Target-only scope

Inventory, validation and batch conversion use recognized direct input folders:
PR, SR, TOS, WHT, TOI, Info, and finance folders/aliases recognized by the system
(GL/General Ledger, TB/Trial Balance, Bank, Sales/Revenue, AR/Receivables,
Expense/Purchase Ledger). Unrelated company-root files and folders are excluded.
Files nested inside a recognized target remain in scope; hidden files/folders
and temporary Excel lock files are ignored.

## Convert without file or destination dialogs

1. Select the company folder. Source overview loads automatically.
2. Select a folder/year row and click **Convert selected row**, or use
   **Convert all target PDFs** to process every recognized target.
3. Progress appears while conversion runs. One failed/scanned PDF does not stop
   the other files. Source overview refreshes automatically when it finishes.
4. Use **Results** to see each converted, already-converted or failed PDF.
   **Open Excel** opens the selected extraction. Double-clicking a converted
   row in validation/results opens Excel; double-clicking a pending PDF retries
   conversion. A selected validation PDF converts only that file.

Original PDFs are retained. New Excel drafts go in a hidden `.pdf-extractions`
folder beside the PDF; Open Excel opens them without having to find that hidden
folder in Finder. Generic drafts stay outside automatic tax import paths, where
they would fail the required V5.1 layout checks. These drafts preserve text,
leading zeros and original page text for review.

A hidden `.tax-audit-conversions.json` in the company root tracks source hashes
and output locations. Keep it and .pdf-extractions with the company folder.
Unchanged PDFs with existing outputs are skipped on rerun. A changed PDF needs
reconversion and gets a new output name; old workbooks are never overwritten.

## Status after conversion

The status changes from **Not ready: PDF needs conversion** to **Converted:
Excel available; review required**. This clears the conversion task rather than
claiming the bank data has been financially validated. Compare balances and all
pages, correct amounts/dates, and map columns before use. Automated financial
reconciliation and OCR are not implemented. Scanned/blank pages remain flagged.
Tax input must still follow the V5.1 Excel headings. A generic PDF extraction is
not automatically a valid PR/SR/TOS/WHT return. No financial-readiness status is
inferred solely from a successful file conversion.

## Tests

28 automated regression tests pass: target scope, menu order, exact PDF table
extraction, preserved sources/outputs, unchanged retry behavior, changed-PDF
reconversion, partial batch failures, automatic path selection, tax draft
isolation, search, caching, original totals and year/month calculations.

A live Tk test on a virtual Linux display passed selected-folder/all-target
conversion, automatic overview refresh, target-only validation, Open Excel from
Results and Bank, partial failures and retry behavior without callback errors.
Native Mac appearance and the contents of your real bank PDFs need your test.

Run `python -m unittest discover -s tests -v` to rerun the automated suite.
