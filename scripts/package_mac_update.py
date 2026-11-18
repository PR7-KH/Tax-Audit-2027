"""Build the downloadable desktop update bundle from the current application files."""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    'control_center.py',
    'engine.py',
    'sales_reconciliation.py',
    'finance_view.py',
    'risk_view.py',
    'heading_filter.py',
    'audit_analysis.py',
    'tos_analysis.py',
    'revenue_expense.py',
    'wht_analysis.py',
    'toi_analysis.py',
    'audit_report.py',
    'ar_ap.py',
    'risk_dashboard.py',
    'tax_paid.py',
    'calc.py',
    'rules/ADT_rules_template.xlsx',
    'scripts/build_adt_template.py',
    'source_check.py',
    'input_tools.py',
    'toi_pdf.py',
    'audit_documents.py',
    'audit_tools_ui.py',
    'tax_law.py',
    'tax_law_ui.py',
    'assets/meeting_note_template.docx',
    'assets/tax_laws.json',
    'requirements.txt',
    'Install_Update_Mac.command',
    'Install_Update_Windows.bat',
    'Start_Control_Center_Windows.bat',
    'setup_windows.py',
)
OUTPUT = ROOT / 'downloads' / 'Tax-Audit-Mac-Update.zip'


def main():
    for name in FILES:
        if not (ROOT / name).is_file():
            raise FileNotFoundError(ROOT / name)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(OUTPUT, 'w', ZIP_DEFLATED) as archive:
        for name in FILES:
            archive.write(ROOT / name, f'Tax-Audit-Mac-Update/{name}')
    with ZipFile(OUTPUT) as archive:
        assert archive.testzip() is None
        assert len(archive.namelist()) == len(FILES)
        for name in FILES:
            assert archive.read(f'Tax-Audit-Mac-Update/{name}') == (ROOT / name).read_bytes()
    print(f'Created and verified {OUTPUT.name}: {len(FILES)} current files')


if __name__ == '__main__':
    main()
