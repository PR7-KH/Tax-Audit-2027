# Tax Audit Engine (Python port of V5.1)

Run (Mac or Windows, Python 3.10+; `pip install -r requirements.txt` once):

    python3 engine.py "<company folder>" "<output>.xlsx"
    python3 engine.py "<company folder>" "<output>.xlsx" --months 3 --link-all

The company folder is the same one you give V5.1: PR / SR / TOS / WHT / TOI / Info.
consolidate.py is the original V5.1 file, unchanged.

Output workbook (all sheets visible, live Excel formulas):
Audit Dashboard, Tax Paid Summary, TOI, FULL TOI, Summary MTR, MTR_PR/SR/TOS/WHT,
PR, SR, TOS, WHT, ADT (rules), Company Info, Log.

Check against a V5.1 workbook built from the same folder (needs LibreOffice):

    python3 -I verify.py output.xlsx Audit_Engine_V5.1_Python.xlsm

Not yet ported: TOI PDF export (Module 10), report builder (Module 11),
the Control Center window (control_center.py).

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

### Download the five-file Mac update

[Download Tax-Audit-Mac-Update.zip](https://github.com/PR7-KH/Tax-Audit-2027/raw/refs/heads/main/downloads/Tax-Audit-Mac-Update.zip)

Unzip it and copy the five files inside `Tax-Audit-Mac-Update` into your existing
app folder, replacing matching files. Keep your templates and company data.
From Terminal in the app folder, run `bash Install_Update_Mac.command`, then
launch normally. Back up your working folder first.

For future updates, rebuild the ZIP with `python3 scripts/package_mac_update.py`
after changing any of the five files, and commit the refreshed ZIP alongside
the source changes.
