"""Build the Windows application with the current Python interpreter."""

import os
from pathlib import Path
import shutil
import subprocess
import sys


ROOT = Path(__file__).resolve().parent


def main():
    if sys.platform != "win32":
        raise SystemExit("Build the Windows executable on Windows.")

    subprocess.run(
        [
            sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
            "--onefile", "--windowed", "--collect-all", "tkinterdnd2",
            "--version-file", "version_info.txt", "--name", "ExcelToiPhoneVCF",
            "src/excel_to_iphone_vcf.py",
        ],
        cwd=ROOT,
        check=True,
    )
    destination = ROOT / "release" / "Excel转通讯录.exe"
    destination.parent.mkdir(exist_ok=True)
    built = ROOT / "dist" / "ExcelToiPhoneVCF.exe"

    certificate = os.environ.get("SIGN_CERT_SHA1")
    if certificate:
        subprocess.run(
            [
                "signtool.exe", "sign", "/sha1", certificate, "/fd", "SHA256",
                "/tr", "http://timestamp.digicert.com", "/td", "SHA256", str(built),
            ],
            check=True,
        )

    # Replace only after a successful build and optional signing step.
    os.replace(built, destination)
    for name in ("build", "dist"):
        shutil.rmtree(ROOT / name)
    (ROOT / "ExcelToiPhoneVCF.spec").unlink(missing_ok=True)
    print(f"Built: {destination.name}")
    if not certificate:
        print("Unsigned build: no code-signing certificate configured.")


if __name__ == "__main__":
    main()
