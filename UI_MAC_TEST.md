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

## TOI PDF export (VBA Module 10 port)

The update ZIP now contains **six** files: copy all of them, including the new
`toi_pdf.py`. In TOIs Overview, select **Export TOI PDF**, enter a step range
(default `1-17`, or for example `4-4`), and click **Export TOI PDFs**. All .xlsx
files directly inside the company's TOI folder are processed. One PDF per input
workbook is saved to `TOI_PDFs`, with progress and per-file results. Existing PDFs
get a numbered successor rather than being replaced. Original workbooks are not
saved or changed. TOI preview/export can run without building the audit dashboard.
The TOI comparison tab is removed from the dashboard preview; its workbook data
is retained internally so dashboard formulas and exports continue to work.

PDF rendering requires LibreOffice. On Mac, install LibreOffice in Applications
from https://www.libreoffice.org/download/download-libreoffice/ and retry; the
app also detects soffice/libreoffice on PATH. Excel is not controlled through
AppleScript. The PDF action gives an explicit setup message when the renderer
is missing. Other application features do not require this installation.

Print columns, row heights, A–D widths, A4 orientation, fit-to-page limits,
centering and margins follow the supplied VBA Module10. E–I AutoFit is
approximated from text width; fonts and pagination can differ between Excel and
LibreOffice. Compare a real company export visually with your trusted VBA PDF.
Nonselected sheets are hidden in a temporary copy. No uploaded VBA is executed.

## Automatic Download-folder setup

Select the company folder containing a direct child named `Download`. Existing Info/TOI folders do not block initial organization. Existing operational
tax/finance folders disable automatic organization; click Organize Download to
explicitly import additional files without replacing existing files.
Otherwise the app moves files with explicit names into newly needed folders:

| Filename token | Destination |
| --- | --- |
| Purchase / PR | PR / year |
| Sale / Sales / SR | SR / year |
| Salary_Tax / TOS | TOS / year |
| WITHHOLDING_TAX / WHT | WHT / year |
| GL / General Ledger | GL |
| TB / Trial Balance | TB |
| Bank / Bank Statement | BANK |
| AR / Receivable | AR |
| Expense / Purchase Ledger | EXPENSE |
| TOI | TOI |
| Info / Company Info | Info |

Monthly tax filenames must identify exactly one 20xx year, for example
`L001_Purchase_01_2025.xlsx`. Missing years, unknown/ambiguous names, unsupported
formats and conflicts are left inside Download and reported under **Import log**.
Files in subfolders of Download are considered, but outside-linked files are
not moved. Source bytes are verified before deletion; existing destinations are
never overwritten. A hidden `.tax-audit-imports.jsonl` records planned/completed
moves and hashes. Keep it for tracing or manually reversing an import. Move
results remain visible until Refresh/company change; the journal remains on disk.
For new files added after target folders exist, use Organize Download. Import log
only displays results; it does not start a transfer. This setup does not infer a file's format
or business correctness from its name; normal validation still applies afterward.

36 regression tests passed, including real LibreOffice PDF rendering/content,
orientation and source/output preservation, Download mappings, ambiguity,
conflicts and existing-target no-op behavior. A live Linux Tk check passed folder
selection/organization, Import log, independent TOI preview, real PDF export,
return to preview and Tax Paid Summary after a partial TOI build.


## Fix for Info + Information + Download layouts

The previous guard treated Info alone as a complete existing target setup and
silently skipped organization. Info and TOI no longer block automatic setup;
Information remains an unrelated folder and is untouched. Select the parent
company folder, not Download itself. Missing/empty Download and automatic-skip
conditions now appear explicitly in the log.

Source overview has an **Organize Download** button. Use it to run an incremental
import even when PR/SR or finance folders already exist. Existing destination
files are preserved; conflicting or unrecognized inputs remain in Download.
Explicit imports show their per-file results automatically. The source actions
wrap into two rows so the button remains visible.

Validated with seven organizer tests, eight preview regressions, and a live Tk
check reproducing the screenshot's folder arrangement, automatic PR creation,
explicit SR import, visible conflict messages and preservation of existing Info.
