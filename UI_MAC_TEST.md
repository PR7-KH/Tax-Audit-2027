# Larger preview and report layout update

## Install on your Mac

Close the app. Back up your working folder, then replace **control_center.py** and
**engine.py** together from this version. Keep your existing consolidate.py,
calc.py, templates, source folders and launchers. Reopen as usual. Previous
versions are recoverable through Git history.

## Interface

- Short header and compact report controls leave more vertical space for data.
- Boxed sidebar entries; Summary MTR and MTR detail are adjacent.
- Source overview has nested File names & formats and Validate tax files entries.
  Inventory lists Excel, PDF and CSV names without importing their contents.
- Fit columns now applies to all tables, including source and finance views, and
  recalculates on resize. Actual widths restores wider content-based columns.
- Expand preview hides the header and sidebar for a taller, full-width report.
  Restore layout brings the menus back without rebuilding the workbook.
- Zoom and 200-row paging remain available. Every worksheet row is accessible.

## Consolidation layout (preview and Excel export)

For two years, each PR/SR/TOS/WHT sheet starts with Grand Total, Year 2025,
Year 2026 (using actual input years), followed by the two source heading rows
and all transactions. Monthly/file subtotal rows and the trailing year summary
are removed from these sheets. MTR detail retains file totals and links. Summary
MTR retains its year figures. Transaction values and numeric totals are unchanged.
The same approach handles one year or more than two years.

## Tax Paid Summary

The former 'earlier year' label was incorrect: the existing month selection
applied to Sheet3. Now Info Sheet2!B3 and Sheet3!B3 identify the actual years;
the older year always sums 12 months and the newer year uses the selected count.
Reversed sheet order is supported. Years must be valid and different; missing
or invalid years are reported rather than silently replaced with 2024/2025.
The control and column headings identify the newer year explicitly.

Opening Tax Paid Summary first reads Info only, without consolidating monthly
returns. Changing its months refreshes that summary in the prepared workbook.
Opening other tax reports builds the full workbook when necessary. Refresh
picks up source-file edits. Export still builds from source, so export timing
is different from cached preview timing.

## Validation and Mac checks

Run `python -m unittest discover -s tests -v` (16 tests). Tests cover transaction
and total preservation through actual Excel export/reopen, retained MTR detail,
normal/reversed years, month counts 1/3/12, summary-only loading, cached workbook
reuse, inventory file formats, paging, styling and column-fit redraw behavior.

Live Tk checks on a virtual Linux display passed normal startup, cached
navigation, zoom, column fit for Source overview, context options, expanded
preview and layout restoration. The preview grew from 1111×613 to 1291×730 in
that test. A 19,000-cell formatting microbenchmark improved from 0.60s to 0.10s;
this does not establish total build time or performance versus Excel VBA.

On your Mac, check the same controls with your company data. Compare grand and
year totals, transaction counts, MTR file totals, and older/newer-year tax paid
amounts against your trusted workbook. Consolidated sheet positions intentionally
changed, so a cell-by-cell layout comparison with the previous export will differ.
Native Mac appearance and actual-company timing remain to be verified.
