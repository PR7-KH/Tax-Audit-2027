"""Consolidate monthly workbooks for the V5.1 Build All Data button."""

import argparse
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from math import fsum
import os
from pathlib import Path
import re
import sys
import stat
from tempfile import NamedTemporaryFile
from xml.etree.ElementTree import iterparse

@dataclass(frozen=True)
class Layout:
    first_row: int
    first_col: int
    last_col: int
    first_heading: str


# Row and column numbers come from the attached V5.1 Modules 2 and 9.
LAYOUTS = {
    "PR": Layout(5, 7, 15, "តម្លៃសរុបលើវិក្កយបត្រ"),
    "SR": Layout(5, 8, 19, "តម្លៃសរុបលើវិក្កយបត្រ"),
    "TOS": Layout(4, 10, 14, "ប្រាក់បៀវត្ស"),
    "WHT": Layout(5, 10, 11, "ទឹកប្រាក់ត្រូវបើក"),
}
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def amount(value):
    # Excel SUM over a range includes numeric cells and ignores text and booleans.
    return value if type(value) in (int, float) else 0


def hidden(path):
    attributes = getattr(path.stat(), "st_file_attributes", 0)
    return path.name.startswith((".", "~$")) or bool(attributes & stat.FILE_ATTRIBUTE_HIDDEN)


GENERIC = Layout(5, 7, 15, "")
MTR_INDEX = "_MTR_File_Rows"
RESERVED = {
    "controls", "log", "summary mtr", "toi", "adt", "audit_dashboard",
    "full toi", "tps_companyinfo", "tax paid summary", MTR_INDEX.casefold(),
    *(f"mtr_{name.lower()}" for name in LAYOUTS),
}
MAX_ROW = 1_048_576


@dataclass
class Source:
    headers: tuple
    header_formats: tuple
    merges: list
    rows: list
    totals: tuple


def sheet_name(folder, used):
    base = re.sub(r"[/\\:?*\[\]]", "_", folder).strip(" '") or "Data"
    name = base[:31].rstrip("'")
    index = 2
    while name.casefold() in used:
        suffix = f"_{index}"
        name = base[:31 - len(suffix)] + suffix
        index += 1
    used.add(name.casefold())
    return name


def read_source(path, layout):
    from openpyxl import load_workbook
    from openpyxl.cell.read_only import ReadOnlyCell
    from openpyxl.utils.cell import get_column_letter, range_boundaries
    from openpyxl.utils.datetime import to_excel
    from openpyxl.worksheet._reader import FORMULA_TAG, VALUE_TAG, WorkSheetParser

    # This reader uses the same cell decoder as openpyxl 3.1.5. Collect
    # header merges, row bounds and required formula caches in the same pass.
    class ConsolidationParser(WorkSheetParser):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.last_row = 0
            self.width = layout.last_col
            self.first_uncached = None

        def parse_cell(self, element):
            cell = super().parse_cell(element)
            row, col = cell["row"], cell["column"]
            has_value = cell["value"] is not None
            formula = element.find(FORMULA_TAG)
            if has_value or formula is not None:
                self.width = max(self.width, col)
                if col == 1:
                    self.last_row = max(self.last_row, row)
            if (
                formula is not None
                and not has_value
                and self.first_uncached is None
                and row >= layout.first_row
                and layout.first_col <= col <= layout.last_col
            ):
                value = element.find(VALUE_TAG)
                if value is None or (value.text is None and element.get("t") != "str"):
                    self.first_uncached = (row, col)
            return cell

    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet = book.worksheets[0]
        header_values = ({}, {})
        header_formats = ({}, {})
        merges = []
        rows = {}
        first_error = None
        with sheet._get_source() as xml:
            parser = ConsolidationParser(
                xml, sheet._shared_strings, data_only=True, epoch=book.epoch,
                date_formats=book._date_formats, timedelta_formats=book._timedelta_formats,
            )
            for _, node in iterparse(xml, events=("end",)):
                if node.tag == NS + "row":
                    row, cells = parser.parse_row(node)
                    node.clear()
                    if row in (2, 3):
                        for cell in cells:
                            col = cell["column"] - 1
                            header_values[row - 2][col] = cell["value"]
                            header_formats[row - 2][col] = ReadOnlyCell(sheet, **cell).number_format
                    elif row >= layout.first_row:
                        values = {}
                        for cell in cells:
                            col, value = cell["column"], cell["value"]
                            if (
                                first_error is None and cell["data_type"] == "e"
                                and layout.first_col <= col <= layout.last_col
                            ):
                                first_error = (row, col, value)
                            if value is None:
                                continue
                            if isinstance(value, (date, time, timedelta)):
                                value = to_excel(value, book.epoch)
                            values[col - 1] = value
                        if values:
                            rows[row] = values
                elif node.tag == NS + "mergeCell":
                    left, top, right, bottom = range_boundaries(node.attrib["ref"])
                    if 2 <= top <= bottom <= 3:
                        merges.append((top - 1, left - 1, bottom - 1, right - 1))
                        parser.width = max(parser.width, right)
                    node.clear()
        last_row, width = parser.last_row, parser.width
        if width > 16_384 or last_row > MAX_ROW:
            raise ValueError("The worksheet exceeds Excel's row or column limit.")
        headers = tuple(tuple(row.get(col) for col in range(width)) for row in header_values)
        formats = tuple(tuple(row.get(col) for col in range(width)) for row in header_formats)
        if layout.first_heading and str(headers[0][layout.first_col - 1]).strip() != layout.first_heading:
            column = get_column_letter(layout.first_col)
            raise ValueError(f"The first worksheet does not have the expected V5.1 heading in {column}2.")
        if parser.first_uncached and parser.first_uncached[0] <= last_row:
            row, col = parser.first_uncached
            raise ValueError(
                f"{get_column_letter(col)}{row} has no saved formula result. "
                "Recalculate and save the source in Excel, then run again."
            )
        if first_error and first_error[0] <= last_row:
            row, col, error = first_error
            raise ValueError(f"{get_column_letter(col)}{row} contains Excel error {error}.")
        if last_row < layout.first_row:
            return Source(headers, formats, merges, [], ())
        values = []
        for row in range(layout.first_row, last_row + 1):
            cells = rows.get(row, {})
            values.append(tuple(cells.get(col) for col in range(width)))
        totals = tuple(
            fsum(amount(row[col]) for row in values)
            for col in range(layout.first_col - 1, layout.last_col)
        )
        return Source(headers, formats, merges, values, totals)
    finally:
        book.close()


def files_in(folder):
    for year in sorted(folder.iterdir(), key=lambda path: path.name.casefold()):
        if not year.is_dir() or hidden(year):
            continue
        for file in sorted(year.iterdir(), key=lambda path: path.name.casefold()):
            if file.is_file() and file.suffix.lower() == ".xlsx" and not hidden(file):
                yield year.name, file


def year_value(year):
    return int(year) if re.fullmatch(r"[1-9][0-9]{3}", year) else year


def write_sheet(book, name, files, layout, styles, log, index_rows):
    sheet = book.add_worksheet(name)
    sheet.freeze_panes(3, 0)
    sheet.set_column(0, 0, 25)
    if layout.first_heading:
        sheet.set_column(1, layout.last_col - 1, 17)
    row = 3
    subtotals = defaultdict(list)
    formats = {}
    first_header = None
    completed = 0
    for year, path in files:
        print(f"{name} / {year} / {path.name}", file=log, flush=True)
        try:
            data = read_source(path, layout)
        except Exception as exc:
            raise ValueError(f"{name} / {year} / {path.name}\n{exc}") from exc
        if first_header is None:
            first_header = data.headers
            merged = {(r, c) for r1, c1, r2, c2 in data.merges
                      for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
            sheet.set_row(1, 28)
            sheet.set_row(2, 26)
            for r, header in enumerate(data.headers, start=1):
                for c, value in enumerate(header):
                    if (r, c) in merged:
                        continue
                    fmt = data.header_formats[r - 1][c] or "General"
                    if fmt not in formats:
                        formats[fmt] = book.add_format({**styles["header_properties"], "num_format": fmt})
                    sheet.write(r, c, value, formats[fmt])
            for r1, c1, r2, c2 in data.merges:
                fmt = data.header_formats[r1 - 1][c1] or "General"
                if fmt not in formats:
                    formats[fmt] = book.add_format({**styles["header_properties"], "num_format": fmt})
                sheet.merge_range(r1, c1, r2, c2, first_header[r1 - 1][c1] or "", formats[fmt])
        if not data.rows:
            print("  No transaction rows", file=log, flush=True)
            continue
        if row + len(data.rows) + len(subtotals) + 5 >= MAX_ROW:
            raise ValueError(f"Consolidation for {name} would exceed Excel's row limit.")
        for values in data.rows:
            sheet.write_row(row, 0, values)
            row += 1
        sheet.write_string(row, 0, path.name, styles["file"])
        year_label = year_value(year)
        sheet.write(row, 1, year_label)
        sheet.write_row(row, layout.first_col - 1, data.totals, styles["amount"])
        subtotals[year].append(data.totals)
        if index_rows is not None and name.upper() in LAYOUTS:
            index_rows.append((name.upper(), row + 1))
        row += 1
        completed += 1
    if not subtotals:
        sheet.write_string(3, 0, "No data found in files")
        return 0
    sheet.write_string(0, 0, f"Consolidated on: {datetime.now():%Y-%m-%d %H:%M:%S}")
    sheet.write_string(0, layout.first_col - 2, "Sum")
    yearly = {
        year: tuple(fsum(col) for col in zip(*totals)) for year, totals in subtotals.items()
    }
    sheet.write_row(0, layout.first_col - 1, tuple(fsum(col) for col in zip(*yearly.values())), styles["amount"])
    row += 2
    sheet.write_string(row, 0, "Year")
    sheet.write_row(row, layout.first_col - 1, first_header[1][layout.first_col - 1:layout.last_col])
    for year, totals in yearly.items():
        row += 1
        sheet.write(row, 0, year_value(year))
        sheet.write_row(row, layout.first_col - 1, totals, styles["amount"])
    return completed


def consolidate(root, output, log, *, excel_job=False):
    import xlsxwriter

    folders = [folder for folder in root.iterdir()
               if folder.is_dir() and folder.name.casefold() != "toi" and not hidden(folder)]
    used = set(RESERVED)
    # Allocate the standard names first so similarly named folders cannot take them.
    folders.sort(key=lambda folder: (folder.name.upper() not in LAYOUTS, folder.name.casefold()))
    sources = [(folder, list(files_in(folder))) for folder in folders]
    if not any(files for _, files in sources):
        raise ValueError("No monthly .xlsx workbooks found. Select the folder above the tax folders and year folders.")
    with NamedTemporaryFile(dir=output.parent, suffix=".tmp", delete=False) as file:
        staged = Path(file.name)
    count = 0
    index_rows = [] if excel_job else None
    try:
        with xlsxwriter.Workbook(staged, {
            "strings_to_formulas": False,
            "strings_to_urls": False,
            "default_date_format": "yyyy-mm-dd",
        }) as book:
            styles = {
                "header_properties": {
                    "font_name": "Khmer OS Siemreap", "font_size": 10, "bold": True,
                    "bg_color": "#D6E4F2", "text_wrap": True, "valign": "vcenter",
                },
                "file": book.add_format({"bold": True, "italic": True}),
                "amount": book.add_format({"num_format": "#,##0"}),
            }
            for folder, files in sources:
                name = sheet_name(folder.name, used)
                count += write_sheet(
                    book, name, files, LAYOUTS.get(folder.name.upper(), GENERIC), styles, log, index_rows,
                )
            if index_rows:
                index = book.add_worksheet(MTR_INDEX)
                index.write_row(0, 0, ("Tax Type", "File Subtotal Row"))
                for row, record in enumerate(index_rows, start=1):
                    index.write_row(row, 0, record)
        if count == 0:
            raise ValueError("None of the monthly workbooks contained transaction rows.")
        os.replace(staged, output)
    finally:
        staged.unlink(missing_ok=True)
    print(f"Consolidated {count} workbooks.", file=log)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--excel-job", action="store_true")
    args = parser.parse_args()
    if args.excel_job:
        folder = args.folder.read_text(encoding="utf-16").rstrip("\r\n")
        log_path = args.output.with_suffix(".log")
        with log_path.open("w", encoding="utf-16") as log:
            try:
                consolidate(Path(folder), args.output, log, excel_job=True)
            except Exception as exc:
                print(f"\n{exc}", file=log)
                return 1
        return 0
    try:
        consolidate(args.folder, args.output, sys.stdout)
    except Exception as exc:
        print(exc, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
