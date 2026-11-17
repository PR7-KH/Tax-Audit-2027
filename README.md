# Tax Audit Engine (Python port of V5.1)

Run (Mac: Python 3.10+; Windows: use the Python 3.12 x64 setup below):

    python3 engine.py "<company folder>" "<output>.xlsx"
    python3 engine.py "<company folder>" "<output>.xlsx" --months 3 --link-all

The company folder is the same one you give V5.1: PR / SR / TOS / WHT / TOI / Info.
consolidate.py is the original V5.1 file, unchanged.

Output workbook (all sheets visible, live Excel formulas):
Audit Dashboard, Tax Paid Summary, TOI, FULL TOI, Summary MTR, MTR_PR/SR/TOS/WHT,
PR, SR, TOS, WHT, ADT (rules), Company Info, Log.

Check against a V5.1 workbook built from the same folder (needs LibreOffice):

    python3 -I verify.py output.xlsx Audit_Engine_V5.1_Python.xlsm

TOI PDF export (Module 10) is available under TOIs Overview → Export TOI PDF
using Microsoft Excel on Mac/Windows, with LibreOffice as an optional alternative.
The desktop Control Center is implemented in control_center.py.
Tax Audit Tools → Document Request reproduces the VBA Module11 checklist using
Info/Template.xlsx. Tax Audit Tools → Meeting Notes saves editable Word documents
with the selected company's Info details.
**Tax Audit Tools → Tax Law Search**, beside Meeting Notes, searches all **741**
entries imported from the supplied V9.5 regulation workbook. Search Khmer or
English article text/references, read full articles, jump between highlighted
matches, filter regulations/languages and copy a reference or article into your
notes. This local search works without Excel or a company folder selected.

Interface updates and Mac upgrade checks: see [UI_MAC_TEST.md](UI_MAC_TEST.md).
Tax Paid Summary uses all 12 months of the older Info year and `--months` of the
newer year. Both year cells (Sheet2!B3 and Sheet3!B3) must be valid and different.
Consolidated tax sheets now put grand/year totals above their transactions;
file-level subtotals remain available in MTR detail.

The latest interface adds all-row consolidated search and local text-PDF
extraction. When updating an existing Mac installation, copy the new
`input_tools.py` and dependency files as well, then run
`bash Install_Update_Mac.command`. See UI_MAC_TEST.md for review requirements;
PDF extraction does not automatically validate tax layouts or bank balances.

### Download the desktop update

[Download Tax-Audit-Mac-Update.zip](https://github.com/PR7-KH/Tax-Audit-2027/raw/refs/heads/main/downloads/Tax-Audit-Mac-Update.zip)

Unzip it and copy all files inside `Tax-Audit-Mac-Update` into your existing
app folder, replacing matching files. Keep your templates and company data.
On Mac, run `bash Install_Update_Mac.command` in Terminal in the app folder,
then launch normally. On Windows, install Python 3.12 x64 and run
Start_Control_Center_Windows.bat; setup runs automatically when needed. Back up your working folder first.

For future updates, rebuild the ZIP with `python3 scripts/package_mac_update.py`
after changing any of the bundled files, and commit the refreshed ZIP alongside
the source changes.

PDF conversion: in **Source inspect**, click **Convert PDFs** (or double-click a ✗ PDF
row). Raw extractions stay under `.pdf-extractions` with source-hash history in
`.tax-audit-conversions.json`; when the PDF holds a real table, a standard Excel with the
**same name as the PDF** is saved next to it and the row turns ✓. See "Source inspect" below.

The update bundle includes twenty-four files, including the Tax Audit Tools/search
modules, `assets/meeting_note_template.docx` and `assets/tax_laws.json`.
Copy the **assets folder** too, and rerun
the Mac installer to install python-docx. Selecting a new
company folder with no recognized target folders automatically organizes explicit
filenames from its Download child. See UI_MAC_TEST.md for routing rules and the
Microsoft Excel setup and Khmer font guidance for TOI PDF export.


If a company has Info/Information plus Download, initial organization now proceeds
normally. In Source inspect, **Organize Download** starts an explicit incremental
import; files it could not place stay in Download and are listed there with the reason.
Existing files are never replaced.
Select the parent company folder containing Download, rather than Download itself.


TOI filenames with year suffixes (`TOI25`, `TOI2X`, `TOI30`, `TOI2025`) now route to TOI.
For Khmer PDF export, choose Microsoft Excel, which uses the same PDF engine as
the old VBA system. Workbook fonts lists source font references; install matching
fonts such as Khmer OS Siemreap and restart Excel. Native Excel still needs your
Mac/Windows test; it is not available in the cloud test machine.


Windows setup uses a separate .venv-win and prebuilt wheels. The uploaded Python
3.15 ARM64 error was a cryptography compiler failure; use Python 3.12 x64 rather
than installing compiler tools. Mac Excel uses stable work copies under TOI_PDFs
to reduce repeated Grant Access prompts; initial OS/Excel approval may still be
needed. MTR preview file names (column B) now open their source workbooks.

TOI PDF export now offers **Save in TOI_PDFs / TOI**. Saving beside the TOI
workbooks may reduce separate folder grants, but Excel can still require access
to the prepared work file and output folder; the app cannot grant that access.
App-generated PDFs saved in TOI are recorded separately from raw source PDFs.

Successful **Build all data** plays a short notification sound. **Refresh**
clears the whole session, including company selection, previews, cached data and
unsaved meeting edits. If a job is running, reset waits for it to finish safely.
Saved documents and source files remain on disk. Tax Audit Tools outputs are saved
under the selected company's **Audit_Tools** folder. Document requests reuse
**Doc_Requirement.xlsx / .pdf**, replacing the previous save of each format.
Meeting notes keep dated, sequential files such as
**Meeting_Note_2026-11-16_Note_1.docx** and **Meeting_Note_2026-11-16_Note_2.docx**.
Their body order is company info, meeting date/time/place, legal references and
the written note; the official heading/title is retained. A large writing pane
sits beside the meeting details, with a separate Requests & follow-up tab.
See [UI_MAC_TEST.md](UI_MAC_TEST.md) for template setup and native testing steps.

The law library preserves the uploaded compilation's references/text and source
order. **Library info** shows the source and its single entry at Excel's text
limit; the stored text is included in full. It is a snapshot of that workbook.
To rebuild the bundled library from a newer compilation without running VBA:

    python3 scripts/import_tax_laws.py "<updated law workbook>.xlsm"

Review the new data, restart the app, run the tests and rebuild the update ZIP.
Tax search adds no new Python dependencies.

### Sales tie-out and risk findings

Control Center → Financial audit → **Tax risk finding** (or
`python3 sales_reconciliation.py "<company folder>" out.xlsx`) compares one year of
sales across the sales listing, GL, the 12 SR returns, the annual TOI and bank
statements, and lists risk findings with severity, amount, evidence and a suggested
audit procedure. Needs Sales, GL and Bank **Excel** files (convert PDFs first); SR and TOI
are used when present. The year is taken from the sales listing; bank USD credits are
converted at 4,000 KHR/USD (`--fx`). **Export this report** saves the workbook where you
choose. Bank cash-in is a reasonableness check, not a tie-out.

### Viewing GL, Sales and Bank data

Control Center → Financial audit → **Sales**, **General ledger**, **Bank** (also Trial
balance, A/R, Expense) now open the actual rows from the Excel/CSV files in that folder.
Use **Find** (all words must match; pick a column to search just that one), the **Month**
list, and click any column heading to sort (click again to reverse). Number columns accept
`>1000000`, `<=500`, `100..200`. A green SUM bar shows the row count and totals of the
filtered rows; **Export this report** saves the filtered rows to Excel. All bank files are
aligned into one list with a **Type** column that classifies each credit (foreign
remittance, in-house transfer, NBC transfer, interest...). A folder with only PDFs still
shows the readiness list.

### Pick a value (account, buyer, bank account)

On Sales, General ledger, Bank, Trial balance, A/R and Expense, choose a **Column**, then a
**Value** from the list of every value in that column (for example Account Name → Cash on
Hand, Account Code → `10201 · Vattanac Bank (USD)`, or one Bank account). Bank opens on the
Account column, so you view and total one account at a time. Month and Find still apply on top.

### Tax risk finding menu

Financial audit → **Tax risk finding** lists, for the tax year of the sales listing:

1. **Total sales**: per monthly tax (SR), per annual return (TOI B0), per sales listing,
   per GL, per trial balance, each with the difference from the GL.
2. **Total cash in and A/R**: cash-in per bank account (compared with the matching GL
   bank account), then the A/R roll-forward (opening, invoiced, collected via bank / cash,
   closing) and the A/R listing.
3. **Total expense**: per PR purchase report, per TOI (B4+B5+B6 and B22), per GL (5xxxx and
   6xxxx, income tax excluded), per trial balance.
4. **Risk findings** with severity, amount, suggested procedure and evidence.

A source that is not in the company folder shows **Not provided**; nothing is guessed. The
Excel copy is saved with **Export this report**. Assumptions: GL expense
accounts start with 5 or 6; TB uses its last Debit and Credit columns; the A/R listing total is
the sum of its Balance (or Outstanding, then Amount) column; PR is compared with GL
accounts-payable credits because purchases include inventory and assets.

### Source inspect and Build all data (simple flow)

1. **Select folder.** Explicitly named files in Download are organized automatically the
   first time.
2. **Source inspect** opens by itself. One line per source and year, with
   ✓ ready · ⚠ usable with a note · ✗ not ready · — not provided (optional).
   It checks the folders *and* reads every file in the same pass (the old Format summary,
   Validate files, Import log, Results and Check source folders screens are merged into it).
   Examples: `✓ SR 2025, 12 months (Jan–Dec)`, `⚠ SR 2024, missing Mar`,
   `✗ WHT 2025, PDF: the monthly return must be the Excel file`, `✗ Bank, PDF: click Convert PDFs`.
   **Show each file** lists every file with its own mark; double-click a row to open it
   (a ✗ PDF row converts that PDF).
3. Fix ✗ rows if you can (**Organize Download**, **Convert PDFs**, or add files and click
   **Refresh**/Source inspect), then click **Build all data**.
4. Build all data builds the tax reports, the first page of every sheet, the GL / Sales /
   Bank tables and the Tax risk finding. Afterwards every menu opens without "Working…".
   Changing the months or the "Fill blank ADT links" option builds again when needed.

Finance files are **ready** when a heading row has an amount column (Amount, Debit / Credit,
Money In / Out, Balance) and, for GL, Bank and Sales, a Date column. A PDF is ready when an
Excel with the same name sits next to it, or when its printed total agrees with a ready
Excel in the folder (a copy). A conversion that produced only text lines, or tables without
an amount column, stays ✗ with the reason: export that report from the accounting system as
Excel, or save a clean Excel with the PDF's name in the same folder. Two files with the same
rows and total are flagged ⚠ so totals are not doubled.

**Consolidated data** now has a small filter row under the search box: Column → Value,
Month, Clear, and a Σ total of the matching rows. Clear returns to the full sheet with its
grand and year totals.

### Excel-style heading filters

In every data view (Sales, General ledger, Bank, Trial balance, A/R, Expense, Consolidated data and
the Audit analysis views) **click a column heading** to open a filter list, like Excel's AutoFilter:
sort A→Z / Z→A, clear the column's filter, search, and tick or untick values (each with its row
count). The search box also takes number conditions: `>1000000`, `<=4000000`, `100..200`. Filters on
several columns combine; a filtered heading shows ⧩, and the green SUM bar lists the active filters.
**Clear** removes all filters. Exported Excel files keep filter buttons on their headings.
In **Consolidated data** the old search box and Filter / Value / Month boxes are gone: filter on the
headings, read the totals in the green **Σ TOTAL** line, and press **⟲ Reset view** to clear every
filter and sort and return to the full sheet with its grand and year totals.

### Audit analysis: Revenue, Expense, TOS (salary tax), WHT (withholding tax)

A new **AUDIT ANALYSIS** group sits below Financial audit. Each menu opens a set of views (buttons
above the table); **Export this report** saves every view to one formatted Excel workbook with
filter buttons and a SUBTOTAL row that follows the Excel filter. Build all data prepares them too.

* **Revenue**: SR monthly returns vs sales listing vs GL sales accounts vs TB vs TOI (B0) for the
  year; month-by-month comparison; invoice match SR ↔ sales listing (only in SR, only in listing,
  amount differs); sales by buyer; SR invoice lines; GL sales lines.
* **Expense**: PR purchases vs GL accounts-payable credits; GL 5xxxx / 6xxxx vs TOI (B4–B6, B22) and
  TB; salary per TOS vs GL salary / wage / payroll accounts; monthly comparison; expense by GL
  account; PR purchases by supplier (flags suppliers not found in GL and missing TIN); PR lines.
* **TOS (salary tax)**: reads every monthly TOS staff list (TOS/<year>/*.xlsx, month from the file
  name such as `TOS_03_2025.xlsx`). Staff are matched by **ID** (by name when the ID is blank).
  Columns are found by heading text (Khmer or English); check **Column mapping** first.
  1. **Foreign staff ≤ 4,000,000 KHR (≈ USD 1,000)**: each month, and one line per person.
  2. **Foreign missing months**: declared, then absent, then declared again (e.g. Jan, Feb, —, Apr).
  3. **Foreign left / joined**: last declared before the newest month on file → request the
     resignation letter for the month after; joiners → request contract and work permit.
     **Monthly headcount** shows staff, Khmer, foreign, joined, left and change each month.
  4. **Khmer missing months**: gaps only; Khmer leavers are not listed.
  Also: **Staff × month** grid (salary per month, MISSING for gaps), **Data checks** (declared twice
  in a month, same ID with different names, same name with different IDs, blank ID / nationality)
  and **All declarations**. Nationality containing Khmer / ខ្មែរ / Cambodia counts as Khmer; a
  blank nationality is checked with foreign staff. A month with no TOS file at all is reported once
  and is not counted as a staff member's missing month.
* **WHT (withholding tax)**: reads every monthly PR and WHT return. Each PR line is classified:
  imports, lines with VAT (VAT-registered supplier), goods / materials, fixed assets and utilities
  are excluded; services, fees, rental, interest and royalty from non-VAT suppliers are WHT-type;
  non-VAT lines whose nature cannot be told are **Unclear** (check the invoice). WHT-type purchases
  are grouped by supplier and month and matched with the WHT return of the same month (or the
  next month, when paid later) by TIN, else by name; each WHT line is used only once. Declared
  items go to **Declared OK**; only **Not declared** / **Partly declared** items are raised in
  **WHT risk** with the estimated WHT (rental 10 %, services / fees / unclear 15 %; confirm
  non-resident 14 %). Also: **Monthly summary**, **Recurring not declared** (same supplier in many
  months, e.g. a landlord), **WHT not in PR**, **PR lines classified** (why each line is or is not
  WHT-type), **WHT declarations** and **Column mapping** (check this first). The nature comes from a
  description column (បរិយាយ / Description) and the supplier name; without them more lines are Unclear.
