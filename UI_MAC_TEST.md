# Finance interface update — Mac test

The update changes only control_center.py. Tax calculations, input readers, report
builders, formulas and Excel exports retain their existing implementation.

## Run

Keep your existing templates and company folders. Replace control_center.py in
your working installation with the new file, then double-click
Start_Control_Center_Mac.command as usual. Alternatively, from the project folder:

```sh
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python control_center.py
```

The app opens fullscreen. Press Escape or use Exit full screen to return to a
normal window. F11 toggles fullscreen (some Macs require Fn+F11).

## Check with your company data

1. Select a company folder and run Check source files.
2. Open Summary MTR, MTR detail, TOI and the dashboard with your existing template.
3. Confirm bold data, alternating row backgrounds, colored sections and totals,
   horizontal/vertical scrolling, and readable Khmer text.
4. Export a report and compare its values/formulas with the previous version.
5. Repeat with another company and check fullscreen exit and normal-window resizing.

Row colors are presentation cues, not additional audit conclusions. Workbook bold
rows are emphasized in previews; source Excel cell-by-cell formatting is not
reproduced exactly by the desktop table. Preview remains limited to 5,000 rows.

## Roll back

Close the app. Copy backup/control_center.py from the downloadable bundle over
control_center.py, then reopen. The bundle's backup is the interface immediately
before this redesign. Keep your own templates, data and exports in place.

## Validation performed in cloud

Passed: Python syntax, row classification, source bold-row emphasis, unchanged
rendered values, visible-row counts, fullscreen-toggle logic and bindings,
four-tax synthetic consolidation, totals and retained formulas, and whitespace
checks. UI checks used mocked widgets: no display server was available. Native
Mac rendering, fullscreen transitions and real-company exports require the manual
checks above. Full dashboard checks need your original ADT template.
