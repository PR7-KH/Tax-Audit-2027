"""Build rules/ADT_rules_template.xlsx (the ADT sheet behind the Audit Dashboard).

    python3 scripts/build_adt_template.py                      # from the ADT sheet of the V5.1 workbook
    python3 scripts/build_adt_template.py --input my_ADT.xlsx  # upgrade your own template

Rules you already wrote (Flag column J / Tax-Audit Issue column K) are kept exactly; only empty rule
cells are filled. Every item row gets Current / Prior links to the TOI (the engine re-points them to
the right TOI row by item code). Column L gets the issue in English. A "Key risk checks" section is
added at the bottom with checks across several TOI lines (profit vs cash, debt over assets, negative
gross profit, expenses over cost ...). All formulas use only IF / AND / OR / N / ABS / MAX / MIN so the
Control Center preview can calculate them without Excel.
"""
import argparse
import re
import sys
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = ROOT / "Audit_Engine_V5.1 Fixed & Fast.xlsm"
OUTPUT = ROOT / "rules" / "ADT_rules_template.xlsx"
CODE = re.compile(r"^[ABCD]\d{1,2}$")

# code: (condition, Khmer issue, English issue).  E[X] current, F[X] prior, G[X] change, H[X] change %
RULES = {
    # ---- current assets
    "A14": ("AND(E[A14]>0,H[A14]>=0.2)", "ស្តុកវត្ថុធាតុដើមកើនឡើងច្រើន ពិនិត្យការរាប់ស្តុក និងបញ្ជីទិញ (PR)",
            "Raw-material stock up 20%+: check stock count and purchase records (PR)."),
    "A15": ("AND(E[A15]>0,OR(H[A15]>=0.2,AND(G[A15]>0,G[B0]<0)))", "ស្តុកទំនិញកើនឡើង ខណៈចំណូលមិនកើន ពិនិត្យការរាប់ស្តុក និងការលក់មិនបានកត់ត្រា",
            "Goods stock up (or up while sales fall): stock count, possible unrecorded sales."),
    "A16": ("AND(E[A16]>0,H[A16]>=0.2)", "ស្តុកផលិតផលសម្រេចកើនឡើង ពិនិត្យការរាប់ស្តុក និងថ្លៃដើមផលិតកម្ម",
            "Finished goods up 20%+: stock count and production cost."),
    "A17": ("AND(E[A17]>0,H[A17]>=0.2)", "ផលិតផលកំពុងផលិតកើនឡើង ពិនិត្យការវាយតម្លៃ",
            "Work in progress up 20%+: check valuation."),
    "A18": ("AND(G[A18]>0,H[A18]>H[B0]+0.2)", "គណនីត្រូវទទួលកើនលឿនជាងចំណូល ពិនិត្យការប្រមូលប្រាក់ និងចំណូលដែលមិនទាន់ប្រកាស",
            "A/R growing faster than revenue: check collections after year end and revenue declared."),
    "A19": ("AND(E[A19]>0,G[A19]>0)", "គណនីត្រូវទទួលផ្សេងៗកើនឡើង អាចជាប្រាក់កម្ចីឱ្យម្ចាស់ភាគហ៊ុន/ភាគីពាក់ព័ន្ធ (ការប្រាក់សន្មត ភាគលាភ)",
            "Other receivables up: may be loans to shareholders / related parties (deemed interest, dividends)."),
    "A20": ("AND(E[A20]>0,H[A20]>=0.5)", "ចំណាយកត់ត្រាមុនកើនឡើងច្រើន ពិនិត្យពន្ធកាត់ទុកលើការជួល និងសេវាបង់មុន",
            "Prepaid expenses up 50%+: check WHT on prepaid rent and services."),
    "A21": ("OR(E[A21]>0.1*E[A0],AND(E[B48]>0,G[A21]+G[A22]<0))", "សាច់ប្រាក់ក្នុងបេឡាច្រើន ឬចំណេញតែសាច់ប្រាក់ថយ ពិនិត្យការរាប់សាច់ប្រាក់ និងចំណូលមិនបានកត់ត្រា",
            "Large cash on hand, or profit made but cash fell: cash count, unrecorded sales or withdrawals."),
    "A22": ("AND(E[B48]>0,G[A21]+G[A22]<0)", "ក្រុមហ៊ុនមានចំណេញ ប៉ុន្តែសាច់ប្រាក់ធនាគារ+បេឡាថយចុះ ពិនិត្យការប្រើប្រាស់ប្រាក់ចំណេញ",
            "Profit made but cash (bank + hand) decreased: trace where the profit went (A/R, stock, assets, loans)."),
    "A23": ("E[A23]>0", "ឥណទានប្រាក់រំដោះពន្ធ ពិនិត្យការគណនាពន្ធអប្បបរមា និងការផ្ទេរ",
            "Prepayment of profit tax credit: check the tax computation and carry-forward."),
    "A24": ("AND(E[A24]>0,G[A24]>0)", "ឥណទានអាករកើនឡើង ពិនិត្យវិក្កយបត្រអាករលើការទិញ (PR) និងការស្នើសុំសងវិញ",
            "VAT credit increasing: check input-VAT invoices (PR) and refund claims."),
    "A25": ("AND(E[A25]>0,G[A25]>0)", "ឥណទានពន្ធដទៃកើនឡើង ពិនិត្យភស្តុតាង", "Other tax credits up: check support."),
    "A26": ("AND(E[A26]>0,H[A26]>=0.5)", "ទ្រព្យសកម្មរយៈពេលខ្លីផ្សេងៗកើនឡើង ពិនិត្យលម្អិត", "Other current assets up 50%+: ask for detail."),
    # ---- equity
    "A29": ("E[A29]<0", "មូលនិធិម្ចាស់ទ្រព្យអវិជ្ជមាន (បំណុលលើសទ្រព្យ) ពិនិត្យប្រភពមូលនិធិ និងភាពបន្តដំណើរការ",
            "Negative equity (debt over assets): going concern and source of funding."),
    "A30": ("G[A30]<>0", "មូលធនផ្លាស់ប្តូរ ពិនិត្យលិខិតក្រសួងពាណិជ្ជកម្ម ភស្តុតាងបង់ប្រាក់ និងពន្ធលើការចុះបញ្ជី",
            "Capital changed: check MOC registration, proof of payment and registration tax."),
    "A31": ("E[A31]>0", "តម្លៃលើសនៃការលក់ភាគហ៊ុន ពិនិត្យកិច្ចសន្យា", "Share premium: check the share agreement."),
    "A33": ("E[A33]>0", "លម្អៀងវាយតម្លៃឡើងវិញ ពិនិត្យមូលដ្ឋានរំលស់ (មិនអាចកាត់កង)", "Revaluation surplus: depreciation on revalued amount is not deductible."),
    "A35": ("E[A35]<0", "ខាតយោងពីមុន ពិនិត្យលក្ខខណ្ឌផ្ទេរខាត", "Accumulated losses: check loss carry-forward conditions."),
    "A36": ("ABS(E[A36]-E[B48])>1000", "ចំណេញក្នុងតារាងតុល្យការ មិនស្មើ B48 ក្នុងតារាងលទ្ធផល",
            "Profit in the balance sheet (A36) differs from the income statement (B48)."),
    # ---- liabilities
    "A38": ("E[A38]>0", "បំណុលភាគីជាប់ទាក់ទិន ពិនិត្យកិច្ចសន្យាខ្ចី ការប្រាក់ ពន្ធកាត់ទុក 14%/15% និងដែនកំណត់ការប្រាក់",
            "Related-party loan: loan agreement, interest, WHT 14%/15%, interest limitation."),
    "A39": ("G[A39]>0", "បំណុលធនាគារកើនឡើង ពិនិត្យកិច្ចសន្យា និងចំណាយការប្រាក់ (B43/B44)",
            "Bank loans up: loan agreement and interest expense (B43/B44)."),
    "A40": ("G[A40]>0", "សំវិធានធនកើនឡើង មិនអាចកាត់កងរហូតដល់បង់ពិតប្រាកដ", "Provisions up: not deductible until paid."),
    "A41": ("G[A41]>0", "បំណុលរយៈពេលវែងផ្សេងៗកើនឡើង ពិនិត្យលម្អិត", "Other long-term liabilities up: ask for detail."),
    "A43": ("E[A43]>0", "ឥណទានវិបារូបន៍ ពិនិត្យការប្រាក់ និងពន្ធកាត់ទុក", "Bank overdraft: check interest and WHT."),
    "A45": ("E[A45]>0", "គណនីត្រូវសងភាគីពាក់ព័ន្ធ ពិនិត្យថ្លៃសេវាគ្រប់គ្រង/សួយសារ ពន្ធកាត់ទុក 14% និងថ្លៃផ្ទេរ",
            "Payables to related parties: management fees / royalties, WHT 14%, transfer pricing."),
    "A46": ("AND(G[A46]>0,H[A46]>=0.2)", "គណនីត្រូវសងផ្សេងៗកើនឡើង ពិនិត្យបំណុលចាស់មិនទាន់សង (ចំណាយក្លែងក្លាយ) និងពន្ធកាត់ទុក",
            "Other payables up 20%+: old unpaid items (fictitious expenses?) and WHT on payment."),
    "A47": ("E[A47]>0", "ចំណូលកត់មុន អាករលើតម្លៃបន្ថែមត្រូវប្រកាសពេលទទួលប្រាក់", "Revenue received in advance: VAT due on receipt."),
    "A48": ("G[A48]>0", "ចំណាយបង្គរកើនឡើង ពិនិត្យពន្ធកាត់ទុកលើសេវា និងការជួលពេលកត់ត្រា", "Accruals up: WHT on accrued services / rent."),
    "A50": ("ABS(E[A50]-E[B47])>1000", "ពន្ធលើប្រាក់ចំណូលត្រូវបង់ មិនស្មើពន្ធក្នុងតារាងលទ្ធផល", "Income tax payable differs from tax expense B47."),
    "A51": ("G[A51]>0", "ពន្ធ-អាករត្រូវបង់កើនឡើង ពិនិត្យការបង់ទាន់ពេល និងការពិន័យ", "Taxes payable up: check timely payment and penalties."),
    # ---- income statement
    "B0": ("H[B0]<=-0.2", "ចំណូលថយចុះច្រើនជាង 20% ពិនិត្យការប្រកាសប្រចាំខែ (SR) និងចំណូលមិនបានប្រកាស",
           "Revenue down 20%+: compare with monthly SR and look for undeclared sales."),
    "B7": ("AND(E[B0]>0,E[B7]<0)", "ចំណេញដុលអវិជ្ជមាន លក់ក្រោមថ្លៃដើម ពិនិត្យថ្លៃផ្ទេរ ស្តុក និងការលក់មិនបានកត់ត្រា",
           "Negative gross profit: selling below cost; transfer pricing, stock, unrecorded sales."),
    "B5": ("AND(E[B5]>0,H[B5]>H[B0]+0.1)", "ថ្លៃដើមទំនិញកើនលឿនជាងចំណូល ពិនិត្យការទិញ (PR) និងស្តុក",
           "Cost of goods growing faster than revenue: check purchases (PR) and stock."),
    "B6": ("AND(E[B6]>0,H[B6]>H[B0]+0.1)", "ថ្លៃដើមសេវាកើនលឿនជាងចំណូល ពិនិត្យលម្អិត", "Cost of services growing faster than revenue."),
    "B9": ("E[B9]>0", "ចំណូលពីការជួល ពិនិត្យពន្ធកាត់ទុក 10% ដោយអ្នកជួល", "Rental income: check 10% WHT withheld by tenants."),
    "B13": ("E[B13]>0", "ឧបត្ថម្ភកធន ជាចំណូលជាប់ពន្ធ", "Subsidies are taxable income."),
    "B15": ("E[B15]>0", "ចំណូលការប្រាក់ ពិនិត្យពន្ធកាត់ទុក 4%/6% ដោយធនាគារ", "Interest income: check WHT 4%/6% by the bank."),
    "B16": ("E[B16]>0", "ចំណេញពីការលក់ទ្រព្យ ពិនិត្យការលក់ (Step 10) និងអាករលើការលក់", "Gain on asset sale: check disposal (step 10) and VAT on the sale."),
    "B22": ("OR(E[B22]>E[B4]+E[B5]+E[B6],E[B22]>E[B0])", "ចំណាយប្រតិបត្តិការធំជាងថ្លៃដើម ឬចំណូល ពិនិត្យចំណាយធំៗ (B23-B41)",
            "Operating expenses bigger than cost of sales or revenue: review large expense lines (B23–B41)."),
    "B23": ("E[B23]>0", "ចំណាយបៀវត្ស ផ្ទៀងផ្ទាត់ជាមួយការប្រកាសពន្ធលើប្រាក់បៀវត្ស (TOS)", "Salary expense: compare with TOS declarations."),
    "B27": ("E[B27]>0", "ចំណាយលើការជួល ពិនិត្យពន្ធកាត់ទុក 10%", "Rental expense: check 10% WHT declared."),
    "B28": ("AND(E[B28]>0,H[B28]>=0.5)", "ចំណាយជួសជុលកើនឡើងច្រើន ពិនិត្យថាជាចំណាយមូលធនឬទេ និងពន្ធកាត់ទុក 15%",
            "Repairs up 50%+: capital or revenue expense; WHT 15% on services."),
    "B29": ("E[B29]>0", "ចំណាយកម្សាន្ត មិនអាចកាត់កង", "Entertainment is not deductible: add back."),
    "B30": ("E[B30]>0", "កម្រៃជើងសា/ផ្សាយពាណិជ្ជកម្ម ពិនិត្យពន្ធកាត់ទុក", "Commissions / advertising: check WHT."),
    "B32": ("E[B32]>0", "អំណោយ កាត់កងបានមិនលើស 5% នៃប្រាក់ចំណេញជាប់ពន្ធ", "Donations: deductible up to 5% of taxable profit."),
    "B33": ("E[B33]>0", "ចំណាយសេវាគ្រប់គ្រង/បច្ចេកទេស ពិនិត្យពន្ធកាត់ទុក 14%/15% និងថ្លៃផ្ទេរ",
            "Management / technical services: WHT 14%/15% and transfer pricing."),
    "B34": ("E[B34]>0", "ចំណាយសួយសារ ពិនិត្យពន្ធកាត់ទុក", "Royalties: check WHT."),
    "B35": ("E[B35]>0", "បំណុលទារមិនបាន ពិនិត្យលក្ខខណ្ឌកាត់កង", "Bad debts: check deductibility conditions."),
    "B36": ("E[B36]>0", "ចំណាយរំលស់ ផ្ទៀងផ្ទាត់ជាមួយរំលស់តាមច្បាប់ពន្ធ (Step 10)", "Depreciation: compare with tax depreciation (step 10)."),
    "B37": ("E[B37]>0", "សំវិធានធន មិនអាចកាត់កង", "Provisions are not deductible: add back."),
    "B40": ("E[B40]>0", "ខាតពីការប្តូរប្រាក់មិនទាន់សម្រេច មិនអាចកាត់កង", "Unrealised exchange loss is not deductible."),
    "B41": ("OR(H[B41]>=0.5,E[B41]>0.1*E[B22])", "ចំណាយផ្សេងៗធំ ឬកើនឡើងច្រើន សុំលម្អិត", "Other expenses large or up 50%+: ask for detail."),
    "B42": ("E[B42]<0", "ខាតពីប្រតិបត្តិការ", "Operating loss."),
    "B43": ("E[B43]>0", "ចំណាយការប្រាក់ ពិនិត្យដែនកំណត់ការប្រាក់ និងពន្ធកាត់ទុក 15%", "Interest expense: interest limitation and WHT 15%."),
    "B44": ("E[B44]>0", "ការប្រាក់បង់ឱ្យអនិវាសនជន ពន្ធកាត់ទុក 14%", "Interest to non-residents: WHT 14%."),
    "B46": ("AND(E[B46]<0,F[B46]<0)", "ខាតពីរឆ្នាំជាប់គ្នា ពិនិត្យប្រភពមូលនិធិ និងភាពត្រឹមត្រូវនៃចំណាយ",
            "Loss in both years: funding source and validity of expenses."),
    "B47": ("AND(E[B0]>0,E[B47]<0.01*E[B0])", "ពន្ធលើប្រាក់ចំណូលតិចជាង 1% នៃផលរបរ ពិនិត្យពន្ធអប្បបរមា",
            "Income tax below 1% of turnover: check minimum tax / prepayment."),
    # ---- cost tables
    "C1": ("ABS(E[C1]-F[A14])>1000", "ស្តុកដើមគ្រា មិនស្មើស្តុកចុងគ្រាឆ្នាំមុន (A14)", "Opening raw-material stock differs from last year's closing (A14)."),
    "C5": ("ABS(E[C5]-E[A14])>1000", "ស្តុកចុងគ្រា មិនស្មើតារាងតុល្យការ (A14)", "Closing raw-material stock differs from the balance sheet (A14)."),
    "C2": ("E[C2]>0", "ការទិញវត្ថុធាតុដើម ផ្ទៀងផ្ទាត់ជាមួយបញ្ជីទិញ (PR)", "Raw-material purchases: compare with PR."),
    "C13": ("E[C13]>0", "សេវាម៉ៅការបន្ត ពិនិត្យពន្ធកាត់ទុក", "Subcontracting services: check WHT."),
    "C15": ("ABS(E[C15]-F[A17])>1000", "ផលិតផលកំពុងផលិតដើមគ្រា មិនស្មើឆ្នាំមុន (A17)", "Opening WIP differs from last year's closing (A17)."),
    "C18": ("ABS(E[C18]-F[A16])>1000", "ស្តុកផលិតផលសម្រេចដើមគ្រា មិនស្មើឆ្នាំមុន (A16)", "Opening finished goods differ from last year's closing (A16)."),
    "C19": ("ABS(E[C19]-E[A16])>1000", "ស្តុកផលិតផលសម្រេចចុងគ្រា មិនស្មើតារាងតុល្យការ (A16)", "Closing finished goods differ from the balance sheet (A16)."),
    "D1": ("ABS(E[D1]-F[A15])>1000", "ស្តុកទំនិញដើមគ្រា មិនស្មើស្តុកចុងគ្រាឆ្នាំមុន (A15)", "Opening goods stock differs from last year's closing (A15)."),
    "D2": ("E[D2]>0", "ការទិញទំនិញ ផ្ទៀងផ្ទាត់ជាមួយបញ្ជីទិញ (PR)", "Purchases of goods: compare with PR."),
    "D5": ("E[D5]>0", "ពន្ធគយនាំចូល ផ្ទៀងផ្ទាត់ជាមួយប្រតិវេទន៍គយ", "Import duties: compare with customs declarations."),
    "D8": ("ABS(E[D8]-E[A15])>1000", "ស្តុកទំនិញចុងគ្រា មិនស្មើតារាងតុល្យការ (A15)", "Closing goods stock differs from the balance sheet (A15)."),
}
# short English for the rows whose Khmer rules were written by hand in V5.1 (A2..A12)
EN_ONLY = {
    "A2": "Land: check the title deed (may be in a private name) and any sale in the year.",
    "A3": "Land improvement: WHT on services.",
    "A4": "Buildings: construction (WHT on services) or check the remaining value.",
    "A5": "Buildings on leased land: construction WHT or remaining value.",
    "A6": "Assets under construction: WHT risk; check the related asset increase.",
    "A7": "Plant and equipment: installation WHT, or check remaining value / disposal.",
    "A8": "Goodwill: price paid vs real value, or disposal.",
    "A10": "Leased assets: check the lease contract (WHT).",
    "A11": "Investments: real value, or sale of shares.",
    "A12": "Other long-term assets: check the asset register.",
}
KEY_CHECKS = [  # (description, current value, compared with, condition, Khmer, English)
    ("Profit made but cash (A21+A22) decreased", "G[A21]+G[A22]", "E[B48]", "AND(E[B48]>0,G[A21]+G[A22]<0)",
     "ក្រុមហ៊ុនមានចំណេញ ប៉ុន្តែសាច់ប្រាក់ថយចុះ ពិនិត្យការប្រើប្រាស់ប្រាក់ចំណេញ (A/R ស្តុក ទ្រព្យ កម្ចី ភាគលាភ)",
     "Profit made but cash fell: trace A/R, stock, asset purchases, loans to related parties, dividends."),
    ("Profit made, A/R and cash not increasing", "G[A18]", "G[A21]+G[A22]", "AND(E[B48]>0,G[A18]<=0,G[A21]+G[A22]<=0)",
     "ចំណេញ មិនឃើញក្នុងសាច់ប្រាក់ ឬគណនីត្រូវទទួល", "Profit is not visible in cash or receivables."),
    ("Debt (A37+A42) greater than assets (A0)", "E[A37]+E[A42]", "E[A0]", "E[A37]+E[A42]>E[A0]",
     "បំណុលលើសទ្រព្យសកម្ម ពិនិត្យប្រភពមូលនិធិ និងភាពបន្តដំណើរការ", "Debt over assets: going concern, who funds the company."),
    ("Gross profit negative (B7)", "E[B7]", "E[B0]", "AND(E[B0]>0,E[B7]<0)",
     "ចំណេញដុលអវិជ្ជមាន លក់ក្រោមថ្លៃដើម", "Negative gross profit: selling below cost."),
    ("Operating expenses (B22) greater than cost of sales", "E[B22]", "E[B4]+E[B5]+E[B6]", "AND(E[B4]+E[B5]+E[B6]>0,E[B22]>E[B4]+E[B5]+E[B6])",
     "ចំណាយប្រតិបត្តិការធំជាងថ្លៃដើមលក់ ពិនិត្យចំណាយធំៗ", "Operating expenses bigger than cost of sales: review large lines."),
    ("Operating expenses greater than revenue", "E[B22]", "E[B0]", "AND(E[B0]>0,E[B22]>E[B0])",
     "ចំណាយប្រតិបត្តិការធំជាងចំណូល", "Operating expenses bigger than revenue."),
    ("Loss in both years (B46)", "E[B46]", "F[B46]", "AND(E[B46]<0,F[B46]<0)",
     "ខាតពីរឆ្នាំជាប់គ្នា", "Loss in N and N-1."),
    ("Loss but cash increased", "G[A21]+G[A22]", "E[B46]", "AND(E[B46]<0,G[A21]+G[A22]>0)",
     "ខាត ប៉ុន្តែសាច់ប្រាក់កើនឡើង ពិនិត្យប្រភពសាច់ប្រាក់", "Loss but cash grew: find the source (loans, capital, unrecorded revenue)."),
    ("Related-party funding increased (A38+A45)", "G[A38]+G[A45]", "F[A38]+F[A45]", "G[A38]+G[A45]>0",
     "ប្រាក់កម្ចី/គណនីត្រូវសងភាគីពាក់ព័ន្ធកើនឡើង ពិនិត្យកិច្ចសន្យា និងពន្ធកាត់ទុក", "Related-party loans / payables up: agreements and interest WHT."),
    ("Stock up while revenue down", "G[A14]+G[A15]+G[A16]+G[A17]", "G[B0]", "AND(G[A14]+G[A15]+G[A16]+G[A17]>0,G[B0]<0)",
     "ស្តុកកើន ខណៈចំណូលថយ", "Inventory up while sales fell."),
    ("Fixed assets bought (change in A1 + depreciation)", "G[A1]+E[B36]+E[C9]+E[C12]", "E[B36]+E[C9]+E[C12]", "G[A1]+E[B36]+E[C9]+E[C12]>1000",
     "មានការទិញទ្រព្យសកម្មរយៈពេលវែង ពិនិត្យ Step 10 អាករ និងពន្ធកាត់ទុកលើសំណង់", "Asset purchases: vouch step 10, input VAT, WHT on construction."),
    ("Balance sheet balances (A0 = A28)", "E[A0]", "E[A28]", "ABS(E[A0]-E[A28])>1000",
     "តារាងតុល្យការមិនស្មើ", "Balance sheet does not balance."),
    ("Debt to assets ratio over 80%", "IF(E[A0]=0,0,(E[A37]+E[A42])/E[A0])", "IF(F[A0]=0,0,(F[A37]+F[A42])/F[A0])",
     "AND(E[A0]>0,E[A37]+E[A42]>0.8*E[A0])", "អនុបាតបំណុលលើទ្រព្យខ្ពស់ជាង 80%", "Debt-to-assets above 80%."),
    ("Gross profit margin fell by 10 points", "IF(E[B0]=0,0,E[B7]/E[B0])", "IF(F[B0]=0,0,F[B7]/F[B0])",
     "AND(E[B0]>0,F[B0]>0,E[B7]*F[B0]<F[B7]*E[B0]-0.1*E[B0]*F[B0])", "អនុបាតចំណេញដុលធ្លាក់ចុះច្រើន", "Gross margin fell by more than 10 points."),
    ("Income tax below 1% of turnover", "E[B47]", "0.01*E[B0]", "AND(E[B0]>0,E[B47]<0.01*E[B0])",
     "ពន្ធលើប្រាក់ចំណូលតិចជាងពន្ធអប្បបរមា 1%", "Income tax below the 1% minimum tax."),
]


def copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def expr(text, rows):
    """E[A21] -> N(E25); F[..] prior; G[..] change; H[..] change % (0 when the prior year is 0)."""
    def sub(m):
        kind, code = m.group(1), m.group(2)
        r = rows.get(code)
        if r is None:
            return "0"
        e, f = f"N(E{r})", f"N(F{r})"
        return {"E": e, "F": f, "G": f"({e}-{f})", "H": f"IF({f}=0,0,({e}-{f})/ABS({f}))"}[kind]
    return re.sub(r"([EFGH])\[([ABCD]\d{1,2})\]", sub, text)


def build(source, output):
    wb_src = load_workbook(source, keep_vba=False)
    src = wb_src["ADT"] if "ADT" in wb_src.sheetnames else wb_src.worksheets[0]
    wb = Workbook()
    ws = wb.active
    ws.title = "ADT"
    for row in src.iter_rows():
        for c in row:
            copy_cell(c, ws.cell(c.row, c.column))
    for k, d in src.column_dimensions.items():
        ws.column_dimensions[k].width = d.width
    for k, d in src.row_dimensions.items():
        if d.height:
            ws.row_dimensions[k].height = d.height
    for m in src.merged_cells.ranges:
        ws.merge_cells(str(m))
    rows = {}
    for r in range(1, ws.max_row + 1):
        code = ws.cell(r, 4).value
        if isinstance(code, str) and CODE.match(code.strip()):
            rows.setdefault(code.strip(), r)
    head = Font(bold=True)
    for r in range(1, ws.max_row + 1):          # heading rows get a column L title
        if str(ws.cell(r, 4).value or "").strip() in ("Items Code", "យោង"):
            ws.cell(r, 12, "Tax/Audit Issue (EN)").font = head
            if not ws.cell(r, 10).value:
                ws.cell(r, 10, "Flag ").font = head
            if not ws.cell(r, 11).value:
                ws.cell(r, 11, "Tax/Audit Issue ").font = head
    added = kept = 0
    for code, r in rows.items():
        for col, letter in ((5, "C"), (6, "D")):
            v = ws.cell(r, col).value
            if v is None or isinstance(v, (int, float)):
                ws.cell(r, col).value = f"=TOI!{letter}{r}"   # re-pointed by item code when the engine builds
        has_rule = ws.cell(r, 10).value not in (None, "") or ws.cell(r, 11).value not in (None, "")
        if code in RULES and not has_rule:
            cond, kh, en = RULES[code]
            c = expr(cond, rows)
            ws.cell(r, 10).value = f'=IF({c},"Check","")'
            ws.cell(r, 11).value = f'=IF({c},"{kh}","")'
            ws.cell(r, 12).value = f'=IF({c},"{en}","")'
            added += 1
        elif has_rule:
            kept += 1
            if code in EN_ONLY and ws.cell(r, 12).value in (None, ""):
                ws.cell(r, 12).value = f'=IF(J{r}="Check","{EN_ONLY[code]}","")'
        for col in (11, 12):
            ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")
    # ---- key risk checks section
    start = ws.max_row + 2
    title_fill = PatternFill("solid", fgColor="1F4E78")
    ws.cell(start, 1, "Section Key Risk Checks (across TOI lines)").font = Font(bold=True, size=12, color="FFFFFF")
    for c in range(1, 13):
        ws.cell(start, c).fill = title_fill
    hdr = start + 1
    for c, text in ((2, "N"), (3, "Check"), (4, "Key"), (5, "Current / value"), (6, "Compared with"),
                    (10, "Flag "), (11, "Tax/Audit Issue "), (12, "Tax/Audit Issue (EN)")):
        ws.cell(hdr, c, text).font = head
        ws.cell(hdr, c).fill = PatternFill("solid", fgColor="BDD7EE")
    for n, (desc, cur, cmp_, cond, kh, en) in enumerate(KEY_CHECKS, start=1):
        r = hdr + n
        c = expr(cond, rows)
        ws.cell(r, 2, n)
        ws.cell(r, 3, desc)
        ws.cell(r, 4, f"KEY{n}")
        ws.cell(r, 5, "=" + expr(cur, rows)).number_format = "0.0%" if "/" in cur else "#,##0"
        ws.cell(r, 6, "=" + expr(cmp_, rows)).number_format = "0.0%" if "/" in cmp_ else "#,##0"
        ws.cell(r, 10, f'=IF({c},"Check","OK")')
        ws.cell(r, 11, f'=IF({c},"{kh}","")')
        ws.cell(r, 12, f'=IF({c},"{en}","")')
        for col in (11, 12):
            ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["L"].width = max(ws.column_dimensions["L"].width or 0, 60)
    ws.column_dimensions["K"].width = max(ws.column_dimensions["K"].width or 0, 60)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return dict(added=added, kept=kept, key_checks=len(KEY_CHECKS), items=len(rows), output=output)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--input", type=Path, default=DEFAULT_SOURCE, help="workbook with the ADT sheet to upgrade")
    p.add_argument("--output", type=Path, default=OUTPUT)
    a = p.parse_args(argv)
    res = build(a.input, a.output)
    print(f"Saved {res['output']}: {res['items']} TOI items, {res['kept']} existing rules kept, "
          f"{res['added']} rules added, {res['key_checks']} key risk checks.")


if __name__ == "__main__":
    sys.exit(main())
