"""Create a separate Windows environment using a wheel-supported Python runtime."""
import json
from pathlib import Path
import re
import subprocess
import sys
from datetime import datetime

ROOT = Path(__file__).resolve().parent
PROBE = "import json,sys,sysconfig; print(json.dumps({'version':list(sys.version_info[:2]),'platform':sysconfig.get_platform(),'executable':sys.executable}))"


def compatible_runtime(info):
    return info.get('version') == [3,12] and info.get('platform') == 'win-amd64'


def runtime_info(command):
    try:
        result = subprocess.run([*command,'-c',PROBE],capture_output=True,text=True,timeout=20)
        if result.returncode:
            return None
        return json.loads(result.stdout.strip())
    except (OSError,subprocess.SubprocessError,ValueError):
        return None


def find_runtime():
    candidates=[('py','-3.12-64'),('py','-3.12'),(sys.executable,),('python',)]
    # A Windows ARM machine may have both ARM and x64 runtimes: probe each path.
    try:
        listed=subprocess.run(['py','-0p'],capture_output=True,text=True,timeout=20)
        for line in listed.stdout.splitlines():
            match=re.search(r'([A-Za-z]:\\.+?python(?:3)?\.exe)\s*$',line,re.I)
            if match:
                candidates.append((match.group(1),))
    except (OSError,subprocess.SubprocessError):
        pass
    for candidate in candidates:
        info=runtime_info(candidate)
        if info and compatible_runtime(info):
            return Path(info['executable'])
    raise RuntimeError('Install Python 3.12 Windows x86-64 (x64), not Python 3.15 ARM64, from python.org/downloads/windows/ and run setup again. Windows on ARM can run x64 Python. No Rust or Visual Studio compiler is needed for the supported runtime.')


def main():
    if sys.platform!='win32':
        raise RuntimeError('Run this setup helper on Windows; use Install_Update_Mac.command on Mac.')
    runtime=find_runtime()
    environment=ROOT/'.venv-win'
    python=environment/'Scripts'/'python.exe'
    existing=runtime_info((str(python),)) if python.is_file() else None
    if environment.exists() and (not existing or not compatible_runtime(existing)):
        backup=ROOT/('.venv-win.backup-'+datetime.now().strftime('%Y%m%d-%H%M%S'))
        index=2
        while backup.exists():
            backup=ROOT/('.venv-win.backup-'+datetime.now().strftime('%Y%m%d-%H%M%S')+f'-{index}');index+=1
        environment.rename(backup)
        print('Preserved incompatible Windows environment as',backup.name)
    if not python.is_file():
        subprocess.run([str(runtime),'-m','venv',str(environment)],check=True)
    # Never fall back to compiling Rust/C packages on the user's computer.
    subprocess.run([str(python),'-m','pip','install','--only-binary=:all:','-r',str(ROOT/'requirements.txt')],check=True)
    subprocess.run([str(python),'-m','pip','check'],check=True)
    print('Windows dependencies ready in .venv-win. The Mac .venv and company files were not changed.')


if __name__=='__main__':
    try:
        main()
    except Exception as exc:
        print('SETUP FAILED:',exc)
        sys.exit(1)
