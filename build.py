import os
from pathlib import Path
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parent
WORK = ROOT / "build"


def run(command):
    print(" ".join(str(part) for part in command), flush=True)
    subprocess.run([str(part) for part in command], cwd=ROOT, check=True)


if not (ROOT / "runtime" / "program" / "soffice.com").is_file():
    raise SystemExit("缺少 runtime 转换组件，无法构建")
with zipfile.ZipFile(ROOT / "runtime.zip", "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
    for source in (ROOT / "runtime").rglob("*"):
        if source.is_file():
            archive.write(source, Path("runtime") / source.relative_to(ROOT / "runtime"))
run([sys.executable, "-m", "unittest", "discover", "-s", ROOT, "-p", "tests.py", "-v"])
run([
    sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile", "--console",
    "--name", "Visio转PDF", "--distpath", ROOT.parent,
    "--version-file", ROOT / "version_info.txt",
    "--workpath", WORK / "work", "--specpath", WORK,
    "--add-data", f"{ROOT / 'runtime.zip'}{os.pathsep}.",
    "--hidden-import", "win32com.client", "--hidden-import", "pythoncom", "--hidden-import", "pywintypes",
    ROOT / "visio2pdf.py",
])
print(f"构建完成：{ROOT.parent / 'Visio转PDF.exe'}")
