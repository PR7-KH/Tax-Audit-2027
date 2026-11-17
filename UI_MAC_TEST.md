# Desktop update: Tax Law Search and Tax Audit Tools

## Install this update

Back up the working app folder. Unzip the update bundle and copy **all nineteen
files, including the assets folder**, into that folder, replacing matching files. Keep templates, company data,
other Python modules and source workbooks in place.

- **Mac:** in Terminal in the app folder, run `bash Install_Update_Mac.command`.
  Then use your existing Start_Control_Center_Mac.command launcher.
- **Windows:** install **Python 3.12 x64 / x86-64** from python.org/downloads/windows/
  if it is not already installed. Then run Start_Control_Center_Windows.bat;
  it invokes setup automatically when needed. Install_Update_Windows.bat can
  also be run directly. Do not choose the ARM64 Python installer for this profile.

Mac uses Python 3.10+ and .venv. Windows uses Python 3.12 x64 and a separate
.venv-win, including on Windows ARM using x64 emulation. Setup
installs the Python Excel bridge (xlwings) and its platform dependencies. No Excel
PDF add-in, xlwings add-in or PRO subscription is required for this feature.
Desktop Microsoft Excel must already be installed and activated for Excel export.
The Tax Audit Tools modules require **python-docx**; rerun the Mac installer even
if the previous version already worked. The Windows launcher checks for this
dependency and runs setup if it is missing.

## TOI filenames and Download organization

Names such as `L001-100121101_TOI25.xlsx`, `TOI2X.xlsx`, `TOI30.xlsx`,
`TOI2025.xlsx`, and `TOI_2026.xlsx` match the TOI folder. Future numeric and
alphanumeric year suffixes following TOI are accepted. TOI workbooks go directly into TOI,
without a year subfolder. Unrelated names containing letters such as STOIC do
not match accidentally.

Select the **company folder containing Download**, not Download itself.
Automatic setup can run when only Info/TOI reference folders exist. Existing
operational tax/finance folders disable automatic setup; Source inspect →
**Organize Download** explicitly imports additional files without replacing
existing files. Files it could not place are listed under Download with the reason.

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
| TOI, TOI25, TOI2025, etc. | TOI |
| Info / Company Info | Info |

Monthly tax filenames must contain exactly one 20xx year. Unknown/ambiguous
names, missing tax years, unsupported extensions and conflicts stay in Download
and are listed under Download in Source inspect. Bytes are verified before deleting a moved source;
`.tax-audit-imports.jsonl` records destinations and hashes. Existing Info,
Information and other company files remain intact.

## TOI PDF export: use Microsoft Excel for Khmer

The uploaded VBA Module10 calls `ExportAsFixedFormat`, which is Excel's own PDF
exporter. It does not use a separate PDF plugin. The app now defaults to
**Microsoft Excel** on Mac and Windows, using a separate Excel instance and
workbook copies. LibreOffice remains an explicit alternative; an Excel
failure does not silently switch to LibreOffice, which can render fonts differently.

1. Open TOIs Overview → Export TOI PDF.
2. Select Microsoft Excel in the renderer dropdown.
3. Enter `1-17` for all steps or a range such as `4-4`.
4. Choose **Save in TOI_PDFs** (default) or **TOI** (beside the source workbooks).
5. Click Export TOI PDFs. All .xlsx files directly in TOI are processed;
   inspect per-file results.

On the first Mac export, allow the launching app (Terminal/Python/your launcher)
to control Microsoft Excel if macOS asks. If denied, change System Settings →
Privacy & Security → Automation. Launch Excel once beforehand to complete its
normal activation/setup. This feature opens only work copies and closes its
own Excel instance. It never saves back to original workbooks. Existing PDFs
receive numbered successors rather than being overwritten.

Mac Excel's **Grant Access** file/folder dialogs are separate from macOS Automation
permission. Earlier exports used a new temporary folder each time, which could
cause repeated grants. Mac Excel now reuses prepared workbook paths under
`<chosen output folder>/.excel-work` and the same staging-PDF directory there. Work
copies are retained between exports so their file identity stays stable. They
contain company data just like the exported PDFs; keep this app-owned work folder
with the company directory. Original workbooks remain unchanged.

Approve Excel's access request for the work file and chosen output folder
when offered. The app cannot bypass sandbox permissions or guarantee Excel never
asks again. It does not display its own input/output file-selection dialogs.
Windows Excel and LibreOffice continue to use temporary work copies.

Choosing TOI may reduce the number of distinct folders Excel needs, but opening
an original workbook directly would lose the prepared print layout and source
protection. Export therefore continues to use stable work copies. Excel may
still ask for both file and folder access on subsequent runs. When PDFs are
saved in TOI, `.tax-audit-toi-exports.json` identifies generated outputs so source
validation does not treat them as raw PDFs needing conversion; keep that file.

Module10 row heights, A–D widths, A4 orientation, print areas, centering, margins
and per-step page limits are retained. Excel additionally performs native E–I
AutoFit. The LibreOffice path approximates AutoFit and can differ visually.

## Khmer fonts

The uploaded audit workbook references **Khmer OS Siemreap**, plus Calibri and
Aptos. Your actual TOI files may reference other fonts. **Workbook fonts** lists
the families referenced by those files; it does not certify that they are
installed on the computer.

On Mac, install the same matching Khmer fonts in **Font Book**, then quit/reopen
Excel and the audit app. On Windows, use Fonts settings if the matching fonts are
missing. Use the working Windows installation's font files or a trusted source.
Do not replace legacy Khmer fonts or transliterate text automatically: the
correct font depends on the original workbook. This update preserves workbook
font names, Khmer text and formulas. Native Excel is the preferred renderer, but
visual correctness still needs verification on your actual computer/document.

If using LibreOffice, select it explicitly and ensure it also has the required
fonts. Its renderer is separate from Excel. LibreOffice is not required for the
Microsoft Excel export option. The older verify.py comparison workflow still
uses LibreOffice independently.

## Target-folder validation, conversion and search

Validation and PDF-to-Excel conversion consider recognized tax/finance folders,
not unrelated company folders. **Convert PDFs** converts every finance PDF that is
not ready yet (double-click a ✗ PDF row for just one). Source inspect refreshes
after completion. When the PDF holds a real table, a standard Excel with the PDF's
name is saved next to it; raw extractions stay under .pdf-extractions. `.tax-audit-conversions.json` tracks source
hashes and outputs, so unchanged PDFs are skipped and originals are preserved.

**Converted: Excel available; review required** clears the conversion task, not
the financial review. Compare every page/balance, correct amounts/dates and map
columns before importing. Scanned PDFs still need OCR. Tax returns still need
V5.1 headings; generic extracted workbooks are excluded from automatic monthly
imports. Financial reconciliation is not implemented by extraction alone.

Consolidated search scans all tax tabs and rows and jumps to matching pages;
200-row paging, cached navigation, zoom and column fitting remain available.
TOI comparison is removed from dashboard tabs but retained internally for its
formulas. MTR summary remains under MTR Overview. Tax Paid Summary uses all
12 months of the older year and selected months of the newer year.

## Validation performed and what needs your test

75 regression tests have passed, including actual LibreOffice PDF rendering,
TOI25 routing, source/output preservation, font/text/formula preservation, Excel
backend choice, stable Mac work-file identity/locks, source hyperlink opening,
Windows runtime rejection of unsupported ARM64/3.15, native API cleanup/COM initialization,
existing totals, chronological years and import conflicts. A live Linux Tk test
passed TOI25 import, workbook-font display, renderer routing/status and clickable
MTR file names. New checks cover Module11 company/item mappings, original-template
preservation, Word field substitution and versioning, generated-output exclusion,
TOI-folder output tracking, build notification routing and deferred reset. A live
Linux Tk check also passed the Tax Audit Tools document tabs, document saves, company
auto-fill, build notification invocation and reset of previews/forms/undo history.
The latest checks verify request replacement (including failed saves/file locks),
text-length underlines, dated/sequential note naming, date validation and document
field order. A live Tk check verified visible selectable list text under white
system text defaults, wrapping/scrolling, repeated saves and the enlarged editor.
The law checks cover case-insensitive literal phrases, code/text matches,
Khmer/English filtering, duplicate references, exact import preservation,
formula rejection, highlight counts, wrapped navigation, clipboard copying,
Khmer helper insertion, read-only display, missing data and clearing pending searches.
All 741 codes/texts and 11 helper characters were also compared directly with the
uploaded source workbook, including its SHA-256. A live app check passed search
without a company folder or launching Excel/browser, company switching and
Refresh; the document/meeting workflow check still passed with the third tab.

**Native Microsoft Excel was mocked in these tests:** cloud has no Mac/Windows
Excel. Real Excel automation, Windows batch launchers, macOS Automation permission
and Khmer shaping require your native test. Dependency wheels were downloaded for
Python 3.10 on Apple Silicon Mac and the complete Python 3.12 Windows x64
profile, but that is not proof of native execution or Mac permission caching. Run `python -m unittest discover -s tests -v` for the suite; the real
LibreOffice test skips when LibreOffice is absent.
The four law UI tests skip if no Tk display is available. The 75-test cloud run
used a virtual display and LibreOffice, with no skips.

To roll back, restore all application files from your backup or Git history.
Your company source workbooks and templates are unchanged by this update.
Saving a document request intentionally replaces its previous app-generated file.

## Tax Audit Tools: Document Request

Open **Tax Audit Tools** below Tax Paid Summary, then **Document Request**.
Put the original **Template.xlsx** in the selected company's **Info** folder.
As in VBA Module11, Sheet1 provides the report layout and Sheet2 provides choices
starting at A7 and ending at the first blank cell (up to row 1000). Keep the
company Info workbook as a separate Excel file in Info with Sheet1.

The list uses explicit dark labels on alternating light rows, so the item names
remain visible with Mac dark-mode text defaults. Click a name or its checkbox;
long names wrap as the window changes size, and the list scrolls with the mouse.

Select up to **17** items. **Save / Open Excel** fills company details in
A6:B10 using Info rows 1, 2, 5, 6, 3, and puts the selected requests in A14:A30.
It hides unused rows 16–30 and preserves the template's page setup.
Accounting underlines are converted to ordinary text underlines, with trailing
padding/underscore decorations removed from underlined headings. Template borders
and blank fill-in lines are preserved, and the source template is unchanged.
**Save / Preview PDF** renders a PDF using the renderer selected in TOIs Overview
(Microsoft Excel on Mac/Windows by default), then opens it for review/printing.
Use the viewer's Print command. Source templates are never overwritten.
Excel and PDF results are saved under **company/Audit_Tools** as
**Doc_Requirement.xlsx** and **Doc_Requirement.pdf**. Repeated saves replace the
same file of that format, without creating numbered copies. Close the file in
Excel/the PDF viewer before saving again if Windows reports it is locked.
Replacement occurs only after the new file is complete; a failed save leaves
the previous file intact. Saving Excel updates the workbook; use Save / Preview
PDF to regenerate the PDF for the current selection. Existing numbered outputs
from older app versions are left on disk.

## Tax Audit Tools: Meeting Notes

The **Meeting Notes** tab starts with the selected company's Info details and
today's Cambodia date in **YYYY-MM-DD** format. The details sit on the left;
the **Writing note** tab on the right provides a large editor. Drag the divider
to change the space allocated to each side. **Requests & follow-up** contains
requested documents, decisions and closing text. Edit details and write the note,
then **Save Word**. Files are kept under **company/Audit_Tools**, for example
**Meeting_Note_2026-11-16_Note_1.docx** then
**Meeting_Note_2026-11-16_Note_2.docx**. The date comes from the meeting-date
field; numbering starts at 1 for each date and does not replace previous notes.
**Open Word / Print** opens the latest saved document in your default Word app.
It opens the saved version, so save again after further edits.

The bundled template is derived from the uploaded meeting-note layout, retaining
its official heading, logo and signature areas. Previous company/case names,
references, dates and findings have been removed. The body now follows the order
**company info → meeting date/time/place → legal/notice references → attendees
and agenda → written note → requests/actions/closing**. Body paragraphs use the
full available page width, with clean spacing and no hanging indents. Blank
optional fields do not leave empty sections. The form does not reuse the old
case's facts. The template references **Khmer MEF1 / Khmer MEF2** fonts; install
the matching fonts for the original appearance. Check the resulting pagination,
Khmer shaping and signature positions in Microsoft Word before printing. Cloud
checks confirm editable text and structural preservation, not native Word layout.

## Tax Audit Tools: Tax Law Search

Open **Tax Audit Tools → Tax Law Search**, directly beside Meeting Notes.
A company folder is optional for law search; document requests and Word saves
still require one. When no company is selected, Tax Audit Tools opens this tab.

Type a reference or phrase in Khmer or English. Results update after a short
typing pause, or press Enter/Search. As in the VBA form, search matches a
**literal, case-insensitive phrase** anywhere in the reference or article text;
results stay in workbook order. Words are not automatically translated or
treated as separate search terms. **Browse all** displays the full library.
The regulation and language filters refine results; the English filter uses
the workbook's `-EN` references.

Select a result to read its full text. Matches are green and the current match
is yellow. **Previous / Next** (or Shift-F3 / F3) wrap through occurrences in
that article. A reference-only match is labelled separately when the phrase
does not appear in the body. The article is read-only, with original newlines;
**+ / -** change reading size. Drag the divider to resize the results/reader.

**Khmer characters** provides the original 11 helper characters and inserts the
chosen character at the search caret. **Copy reference / Copy article** puts
the selected material on the clipboard for pasting into Meeting Notes or another
document. Ordinary Ctrl/Cmd-F focuses the search box while working in this tab.
**Clear** resets this search. Navigation away/back and changing companies retain
the law search; the global **Refresh** clears it with the rest of the session.

The bundled `assets/tax_laws.json` contains **741 entries** (378 Khmer-reference
entries and 363 `-EN` entries) from **Advance Regulation Compiled - V9.5 Pro-Alt
Fixed .xlsm**. Codes, text, source rows and the character helper are preserved.
The app uses the local JSON library at runtime. Neither Excel/VBA nor the
original Windows browser control is required for searching or reading.

**Library info** identifies the source compilation. One source cell, row 378
(`574 MEF(01-Jan-25)(TP) 1-13`), contains 32,767 characters, Excel's cell limit.
The library retains every character stored there; it cannot recover material
not present in the supplied workbook. Law wording/currency follows that uploaded
compilation and is not refreshed from the internet.

For a newer compilation, the development helper
`python3 scripts/import_tax_laws.py "<workbook>.xlsm"` rebuilds the JSON from
Regulations A/B and the optional Khmer sheet, without executing VBA. It validates
rows and rejects formula/incomplete entries instead of silently omitting them.
Review and test the replacement library, restart the app and rebuild the update
ZIP before distributing it. The original workbook is never changed by import.

## Build notification and Refresh

A successful **Build all data** plays a short system notification: Glass on Mac,
the Windows notification sound, or a Tk bell elsewhere. Playback follows system
volume/settings; native playback still needs your test. Failed builds do not
play the completion sound.

**Refresh** now returns the interface to startup: company selection, built data,
previews, search, report selection, zoom and unsaved meeting edits are cleared.
Save Word before refreshing. Saved outputs and source files remain on disk.
If a task is running, Refresh waits for it to finish safely, then clears the
session and discards its UI result. It does not interrupt a file write/export.


## Windows installation log diagnosis

The reported error is CPython 3.15 ARM64 attempting to compile cryptography and
failing because link.exe/MSVC is missing. Setup now selects a supported Python
3.12 x64 runtime and installs prebuilt wheels only. It does not install Rust,
Visual Studio or build dependencies. If this runtime is missing, it reports the
exact required installer rather than attempting compilation. Existing incompatible
.venv-win folders are backed up before rebuilding. The shared Mac .venv is not
changed, which matters when the project is shared between Mac and Windows.

## Open MTR sources

In MTR_PR, MTR_SR, MTR_TOS or MTR_WHT, **click the file name in column B** to open
its original Excel workbook. Linked rows are blue and the filename has a hand
cursor. This works on cached/paged previews. The Excel export also retains file
hyperlinks, using absolute file URIs with spaces/Unicode encoded and the actual
source folder's casing. A missing source produces a status message; data/totals
are not changed.

## What the import controls do

**Organize Download** starts a file transfer into the appropriate source folders.
Source inspect then shows what is still in Download and why (for example an
unrecognized file name). The on-disk .tax-audit-imports.jsonl remains the
persistent transfer record.
