"""Build the downloadable desktop update bundle from the current application files."""
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    'control_center.py',
    'engine.py',
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
