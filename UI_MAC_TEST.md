# Desktop update: Excel PDF export and TOI filenames

## Install this update

Back up the working app folder. Unzip the update bundle and copy **all eight
files** into that folder, replacing matching files. Keep templates, company data,
other Python modules and source workbooks in place.

- **Mac:** in Terminal in the app folder, run `bash Install_Update_Mac.command`.
  Then use your existing Start_Control_Center_Mac.command launcher.
- **Windows:** run Install_Update_Windows.bat once, then launch with
  Start_Control_Center_Windows.bat.

Python 3.10+ is required. Each computer needs its own virtual environment. Setup
installs the Python Excel bridge (xlwings) and its platform dependencies. No Excel
PDF add-in, xlwings add-in or PRO subscription is required for this feature.
Desktop Microsoft Excel must already be installed and activated for Excel export.

## TOI filenames and Download organization

Names such as `L001-100121101_TOI25.xlsx`, `TOI26.xlsx`, `TOI2025.xlsx`, and
`TOI_2026.xlsx` now match the TOI folder. TOI workbooks go directly into TOI,
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
temporary workbook copies. LibreOffice remains an explicit alternative; an Excel
failure does not silently switch to LibreOffice, which can render fonts differently.

1. Open TOIs Overview → Export TOI PDF.
2. Select Microsoft Excel in the renderer dropdown.
3. Enter `1-17` for all steps or a range such as `4-4`.
4. Click Export TOI PDFs. All .xlsx files directly in TOI are processed.
5. Inspect per-file results; PDFs are saved in TOI_PDFs.

On the first Mac export, allow the launching app (Terminal/Python/your launcher)
to control Microsoft Excel if macOS asks. If denied, change System Settings →
Privacy & Security → Automation. Launch Excel once beforehand to complete its
normal activation/setup. This feature opens only temporary copies and closes its
own Excel instance. It never saves back to original workbooks. Existing PDFs
receive numbered successors rather than being overwritten.

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

46 regression tests have passed, including actual LibreOffice PDF rendering,
TOI25 routing, source/output preservation, font/text/formula preservation, Excel
backend choice, temporary-copy handling, native API cleanup/COM initialization,
existing totals, chronological years and import conflicts. A live Linux Tk test
passed TOI25 import, workbook-font display and renderer routing/status.

**Native Microsoft Excel was mocked in these tests:** cloud has no Mac/Windows
Excel. Real Excel automation, Windows batch launchers, macOS Automation permission
and Khmer shaping require your native test. Dependency wheels were downloaded for
Python 3.10 on Apple Silicon Mac and Windows x64, but that is not proof of native
execution. Run `python -m unittest discover -s tests -v` for the suite; the real
LibreOffice test skips when LibreOffice is absent.

To roll back, restore all application files from your backup or Git history.
Your company workbooks and templates are unchanged by this update.
