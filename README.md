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
The VBA report builder (Module 11) remains separate from this workflow.

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

Batch conversion: select a source folder/year row and choose **Convert selected
row**, or use **Convert all target PDFs**. Validation/conversion exclude unrelated
folders. Outputs are preserved under `.pdf-extractions`, with source-hash history
in `.tax-audit-conversions.json`; use **Results → Open Excel** to inspect them.
The overview refreshes automatically and distinguishes conversion completion from
reviewed/import-ready financial data. See UI_MAC_TEST.md for the full workflow.

The update bundle includes `toi_pdf.py` and Mac/Windows setup helpers (nine files total). Selecting a new
company folder with no recognized target folders automatically organizes explicit
filenames from its Download child. See UI_MAC_TEST.md for routing rules and the
Microsoft Excel setup and Khmer font guidance for TOI PDF export.


If a company has Info/Information plus Download, initial organization now proceeds
normally. In Source overview, **Organize Download** starts an explicit incremental
import; **Import log** only displays results. Existing files are never replaced.
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
