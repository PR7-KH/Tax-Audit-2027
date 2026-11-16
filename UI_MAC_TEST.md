# Desktop update: Audit Tools, session reset and PDF export

## Install this update

Back up the working app folder. Unzip the update bundle and copy **all twelve
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
The new Audit Tools modules require **python-docx**; rerun the Mac installer even
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
operational tax/finance folders disable automatic setup; Source overview →
**Organize Download** explicitly imports additional files without replacing
existing files. Import log only displays results.

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
and appear in Import log. Bytes are verified before deleting a moved source;
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
not unrelated company folders. Convert selected row processes a folder/year or
individual PDF; Convert all target PDFs processes every recognized target.
Source overview refreshes after completion. Results → Open Excel opens extracted
workbooks under .pdf-extractions. `.tax-audit-conversions.json` tracks source
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

63 regression tests have passed, including actual LibreOffice PDF rendering,
TOI25 routing, source/output preservation, font/text/formula preservation, Excel
backend choice, stable Mac work-file identity/locks, source hyperlink opening,
Windows runtime rejection of unsupported ARM64/3.15, native API cleanup/COM initialization,
existing totals, chronological years and import conflicts. A live Linux Tk test
passed TOI25 import, workbook-font display, renderer routing/status and clickable
MTR file names. New checks cover Module11 company/item mappings, original-template
preservation, Word field substitution and versioning, generated-output exclusion,
TOI-folder output tracking, build notification routing and deferred reset. A live
Linux Tk check also passed both Audit Tools tabs, document saves, company
auto-fill, build notification invocation and reset of previews/forms/undo history.

**Native Microsoft Excel was mocked in these tests:** cloud has no Mac/Windows
Excel. Real Excel automation, Windows batch launchers, macOS Automation permission
and Khmer shaping require your native test. Dependency wheels were downloaded for
Python 3.10 on Apple Silicon Mac and the complete Python 3.12 Windows x64
profile, but that is not proof of native execution or Mac permission caching. Run `python -m unittest discover -s tests -v` for the suite; the real
LibreOffice test skips when LibreOffice is absent.

To roll back, restore all application files from your backup or Git history.
Your company workbooks and templates are unchanged by this update.

## Audit Tools: Document Request

Open **Audit Tools** below Tax Paid Summary, then **Document Request**.
Put the original **Template.xlsx** in the selected company's **Info** folder.
As in VBA Module11, Sheet1 provides the report layout and Sheet2 provides choices
starting at A7 and ending at the first blank cell (up to row 1000). Keep the
company Info workbook as a separate Excel file in Info with Sheet1.

Select up to **17** items. **Create / Open Excel** fills company details in
A6:B10 using Info rows 1, 2, 5, 6, 3, and puts the selected requests in A14:A30.
It hides unused rows 16–30 and preserves the template's page setup.
**Preview / Print PDF** renders a PDF using the renderer selected in TOIs Overview
(Microsoft Excel on Mac/Windows by default), then opens it for review/printing.
Use the viewer's Print command. Source templates are never overwritten.
Excel and PDF results are saved under **company/Audit_Tools** with numbered
versions. The Excel report remains available if PDF rendering fails.

## Audit Tools: Meeting Notes

The **Meeting Notes** tab starts with the selected company's Info details and
today's Cambodia date. Edit those details and enter attendees, reference, agenda,
discussion, requested documents, decisions and closing text. **Save Word** creates
**company/Audit_Tools/Meeting_Note.docx**; later saves create numbered versions.
**Open Word / Print** opens the latest saved document in your default Word app.
It opens the saved version, so save again after further edits.

The bundled template is derived from the uploaded meeting-note layout, retaining
its official heading, logo and signature areas. Previous company/case names,
references, dates and findings have been removed. The form does not reuse that
case's facts. The template references **Khmer MEF1 / Khmer MEF2** fonts; install
the matching fonts for the original appearance. Check the resulting pagination,
Khmer shaping and signature positions in Microsoft Word before printing. Cloud
checks confirm editable text and structural preservation, not native Word layout.

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
**Import log** only displays the latest moved/skipped/error results. Clicking the
log does not run another transfer. The on-disk .tax-audit-imports.jsonl remains the
persistent transfer record.
