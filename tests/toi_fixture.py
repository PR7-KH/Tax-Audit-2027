"""A small but complete TOI workbook for tests: balanced balance sheets for N and N-1, an income
statement, the cost-of-goods table and a step 10 asset table. The company makes a profit while its
cash falls, holds more stock and receivables, buys plant and pays tax below 1 % of turnover."""
from pathlib import Path

from openpyxl import Workbook

# code: (current N, prior N-1)
VALUES = {
    "A0": (5800, 4900), "A1": (3600, 3000), "A2": (1000, 1000), "A7": (2600, 2000),
    "A13": (2200, 1900), "A15": (900, 500), "A18": (800, 400), "A21": (50, 100), "A22": (450, 900),
    "A28": (5800, 4900), "A29": (1080, 700), "A30": (1000, 1000), "A35": (-300, -200), "A36": (380, -100),
    "A37": (3200, 3000), "A38": (3200, 3000), "A42": (1520, 1200), "A46": (1520, 1200),
    "B0": (5000, 4000), "B2": (5000, 4000), "B5": (3000, 2600), "B7": (2000, 1400), "B22": (1600, 1500),
    "B23": (600, 600), "B27": (200, 200), "B36": (300, 250), "B41": (500, 450), "B42": (400, -100),
    "B46": (400, -100), "B47": (20, 0), "B48": (380, -100),
    "D1": (500, 450), "D2": (3400, 2650), "D7": (3900, 3100), "D8": (900, 500), "D9": (3000, 2600),
}
STEPS = {4: [f"A{i}" for i in range(0, 53)], 5: [f"B{i}" for i in range(0, 49)],
         6: [f"C{i}" for i in range(1, 21)], 7: [f"D{i}" for i in range(1, 10)]}


SCALE = 1_000_000   # amounts in KHR: 5,000 above means 5,000,000,000 riel


def make_toi(root, values=None, step10=True, name="TOI2025.xlsx"):
    values = {k: (v[0] * SCALE if v[0] is not None else None, v[1] * SCALE if v[1] is not None else None)
              for k, v in dict(VALUES, **(values or {})).items()}
    d = Path(root) / "TOI"
    d.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); wb.remove(wb.active)
    for n in range(1, 18):
        ws = wb.create_sheet(f"step{n}")
        ws["A1"] = f"TOI step {n}"
        if n in STEPS:
            ws.append([None]); ws.append([None])
            ws.append(["បរិយាយ", "យោង", "ការិយបរិច្ឆេទនេះ (N)", "ការិយបរិច្ឆេទមុន (N-1)"])
            for code in STEPS[n]:
                cur, pri = values.get(code, (None, None))
                ws.append([f"Item {code}", code, cur, pri])
        if n == 10 and step10:
            ws.append(["ប្រភេទទ្រព្យសកម្ម", "តម្លៃដើមគ្រា", "ការទិញ ឬបន្ថែមក្នុងគ្រា", "ការលក់ចេញ", "តម្លៃចុងគ្រា"])
            for row in (["Plant and machinery", 2500, 900, 0, 3400], ["Land", 1000, 0, 0, 1000], ["សរុប", 3500, 900, 0, 4400]):
                ws.append([row[0]] + [v * SCALE for v in row[1:]])
    path = d / name
    wb.save(path)
    return path


def make_info(root):
    d = Path(root) / "Info"
    d.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    for row in (["Company", "DEMO CO., LTD"], ["TIN", "K001-000000001"], ["Address", "Phnom Penh"], ["", ""],
                ["Tax year", "2025"], ["Activity", "Trading"]):
        ws.append(row)
    wb.save(d / "Info.xlsx")
