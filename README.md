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
