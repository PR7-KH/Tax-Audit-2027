# Source workflow, conversion and search update

## Mac installation

Close the app and back up the working folder. Copy these files from the latest
GitHub version into that folder **together**:

- control_center.py
- engine.py
- input_tools.py (new)
- requirements.txt
- Install_Update_Mac.command (new)

Keep your original templates, consolidate.py, calc.py, company folders and
exports. In Terminal, change to this app folder and run:

```sh
bash Install_Update_Mac.command
```

This installs the pinned PDF dependency into .venv. Then launch normally with
Start_Control_Center_Mac.command. Python 3.10+ is required. Source loading and
report viewing remain local; PDF contents are not sent to a service.

## Layout

The sidebar is wider. Source overview and Build all data are now beside Select
folder, rather than in the report sidebar. Source overview offers Format summary,
Validate files and Convert PDF to Excel. Its summary shows counts such as
`9 files / 6 Excel, 3 PDF`, without listing file names.

MTR detail tabs are MTR_PR, MTR_SR, MTR_TOS, MTR_WHT and Summary MTR. TOIs Overview
contains TOI and FULL TOI. Tax Paid Summary no longer has a Company Info tab;
Company Info remains available to the dashboard and full workbook where needed.

## Readiness and conversion

File presence is not an import-readiness claim. Inventory flags PDFs, legacy XLS,
unsupported tax formats and monthly tax files outside year folders. Validate
files runs the existing monthly tax reader and adds all-folder format/readability
checks, including GL and Bank. Readable finance files still require column mapping
and reconciliation rules, which are not implemented by this update.

Convert PDF to Excel prompts for a PDF and a **new** output workbook. It extracts
text-based tables plus a page-text copy and adds a Conversion Review sheet. It
keeps values as literal text to preserve leading zeros and avoid inventing date,
amount or formula interpretations. The original PDF and existing outputs are
never overwritten. Scanned/image-only, blank or unreadable pages are rejected;
OCR is not included. Text-only fallback pages require manual table reconstruction.

Before importing extracted data, compare every page, opening/closing balances,
transaction counts and totals with the PDF. Correct dates/amounts, assign numeric
columns and map to the required input layout. Tax returns must still follow the
V5.1 Excel headings. A generic extracted bank sheet does not implement automated
bank reconciliation. Validation marks extracted workbooks for review.

Keep conversion drafts and original PDFs outside tax import directories. After
review, place only validated/mapped input workbooks there. Original PDFs can be
archived outside the selected import folders; Refresh then updates inventory.
Keeping a PDF in an import folder intentionally retains its conversion warning.

## Search

Consolidated data has a search box. Find scans all PR/SR/TOS/WHT tabs and all rows,
including rows not loaded into the current preview. Results jump to and highlight
the matching row; Next match cycles through matching rows. This retains fast
200-row paging rather than loading all rows into the visible table. Search and
preview navigation do not modify worksheets or exports.

## Validation and manual checks

Automated checks cover format counts, PDF readiness, actual table extraction,
literal values and leading zeros, original-file preservation, refusal to replace
outputs, blank/scanned-page rejection, search beyond row 6,000, grouped tabs,
existing totals/transactions, tax-year/month handling and cached navigation.
A live Tk test on a virtual Linux display also exercised the conversion dialog
and worker, source validation, search highlighting across tax sheets, grouped
tabs and layout restoration without callback errors.

All 23 tests passed. Dependency wheels were resolved and downloaded for Python
3.10 on both Apple Silicon and Intel Mac targets (this does not establish native
Mac execution). Run `python -m unittest discover -s tests -v`. On your Mac, check menu sizing,
real bank PDF extraction, search on your largest company and numeric totals
against your trusted source. Generic extraction is not certified for every bank
layout. Native Mac behavior and actual bank statements still need your review.

To roll back, restore the previous control_center.py and engine.py together from
your backup or Git history. No source workbooks were changed by the update.
