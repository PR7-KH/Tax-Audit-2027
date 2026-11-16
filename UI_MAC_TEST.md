# Compact workspace and responsive preview — Mac test

Replace **control_center.py** in your existing working app folder. Keep your
existing engine.py, calc.py, consolidate.py, templates and company data. Launch as
usual. The tax calculation and export implementations are unchanged.

## Interface

- Opens maximized with the normal Mac title bar, rather than native fullscreen.
- Compact company information gives the report most of the window height.
- Selecting a folder automatically lists PR/SR/TOS/WHT file counts by year,
  TOI/Info counts and financial-folder counts, without opening workbooks.
- Source overview is an inventory, not validation of workbook contents. Use
  **Validate tax files** when you want the existing detailed source validation.
- MTR detail uses sheet buttons; duplicate sidebar entries are removed.
- The earlier-year month setting appears only in Tax Paid Summary; click Apply
  after editing it. This changes the report options and requires a new build.
- Preview uses 200-row pages. Next/Previous reaches every row, including rows
  after 5,000. Revisiting cached pages does not calculate them again.
- Use −/+ for 50–150% zoom. Fit columns toggles between compact equal widths and
  content-based widths; horizontal scrolling is available for long cells.

## Cache behavior

Build all data prepares the workbook once. Report navigation reuses it, and
previews evaluate only the selected sheet's page. Up to 32 pages are cached;
older pages can be reformatted after eviction, without reconsolidation. Source
file edits are picked up with Refresh, folder selection, or a changed build
option. Export still rebuilds from source, as in the existing application.

## Check on your Mac

1. Select a company and confirm PR counts for each year and the GL file count.
2. Build all data; open Consolidated data, switch reports and return. Check that
   the displayed data is ready without another consolidation.
3. Visit the last preview page in a sheet with more than 5,000 rows. Compare the
   row numbers and totals against your existing workbook.
4. Test zoom, column fitting, scrolling, MTR sheet buttons and normal resizing.
5. Open Tax Paid Summary, change months and Apply; compare results and exports
   with your previous version using the same settings and source files.
6. Select another company, then edit a source file and Refresh. Confirm that
   previous-company and stale-page values are not reused.

## Validation

Run `python -m unittest discover -s tests -v`. Regression checks cover inventory
counts without opening workbooks, automatic inventory, paging beyond row 5,000,
formula-value preservation, cached navigation, workbook reuse, option changes,
normal-window maximization and prevention of simultaneous export jobs. A live Tk
window test on a virtual Linux display also passed startup, automatic inventory,
200-row paging beyond row 5,000, cached return visits, zoom, column fitting and
contextual month controls with no callback errors. For a synthetic 8,201-row,
two-column workbook, first preview navigation took 0.14s and cached return took
0.13s; the workbook was built once. These are preview timings, not a real-company
build benchmark. Native
Mac appearance and timing with your actual workbooks require manual testing.

## Rollback

Close the app and restore control_center.py from the previous GitHub commit or
your local backup. No calculation modules or source workbooks were changed.
