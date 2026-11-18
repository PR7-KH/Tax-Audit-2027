# Sample company folder (anonymous)

`Company XYZ/` is a real company folder with **one file per target folder**, made anonymous with
`python3 scripts/make_sample.py "<real company folder>" "samples/Company XYZ"`:
company, supplier, buyer, payee, employee and shareholder names are replaced (Supplier 001, Employee 0001 …),
TIN / ID-card / passport numbers are fake numbers of the same format, government payees (customs, NSSF …)
are kept. Layouts, headings, Khmer type texts and amounts are the real ones, so the app can be developed and
tested against the real GDT formats.

| Folder | File | Layout |
|---|---|---|
| PR/2025 | `L001-000000000_PURCHASE_01_2025.xlsx` | GDT purchase journal: title row 1, headings rows 2–3 (supplier type / TIN / name in D–F), total row 4, data from row 5, description in P |
| SR/2025 | `L001-000000000_SALE_01_2025.xlsx` | GDT sales journal: buyer type / TIN / name in E–G, amount H, description V |
| TOS/2025 | `L001-000000000_SALARY_TAX_01_2025.xlsx` | GDT salary tax list: personal tax ID B, ID card / passport C, name D, nationality E, salary J, FB K, tax L, FB tax M. Large companies file one month in parts: `..._01_2025.xlsx`, `... (1).xlsx`, `... (2).xlsx` |
| WHT/2025 | `L001-000000000_WITHHOLDING_TAX_01_2025.xlsx` | GDT WHT list: invoice C, payee type / ID / name E–G, transaction type H, rate I, amount J, WHT K |
| TOI | `TOI-L001-000000000_2026-04-20.xlsx` | Annual return, sheets Step1–Step17; the file date is the filing date, the tax year is in Step 1 |
| GL | `General_Ledger_2025_01_sample.xlsx` | GL Transactions (January only): Account Code, Account Name, Date, Period, Transaction Type, Reference No., Name / Counterparty, Memo, Split Code, Split Account, Debit, Credit, Running Balance |

Not included yet: Bank, A/R, A/P (PDF only for now), Info, Tax Paid, Audit Report.
