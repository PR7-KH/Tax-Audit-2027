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

Batch conversion: select a source folder/year row and choose **Convert selected
row**, or use **Convert all target PDFs**. Validation/conversion exclude unrelated
folders. Outputs are preserved under `.pdf-extractions`, with source-hash history
in `.tax-audit-conversions.json`; use **Results → Open Excel** to inspect them.
The overview refreshes automatically and distinguishes conversion completion from
reviewed/import-ready financial data. See UI_MAC_TEST.md for the full workflow.

The update bundle includes fifteen files, including the Tax Audit Tools/search
modules, `assets/meeting_note_template.docx` and `assets/tax_laws.json`.
Copy the **assets folder** too, and rerun
the Mac installer to install python-docx. Selecting a new
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
