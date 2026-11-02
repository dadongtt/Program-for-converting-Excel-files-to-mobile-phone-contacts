"""Package an existing Windows build with example data and dependency licenses."""

import hashlib
import importlib.metadata as metadata
from pathlib import Path
import shutil
import sys
import tkinter
import urllib.request
import zipfile


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from excel_to_iphone_vcf import APP_VERSION


def copy_notices(bundle):
    notices = bundle / "third-party-licenses"
    notices.mkdir(exist_ok=True)
    inventory = [f"Python {sys.version.split()[0]}"]
    packages = [
        "openpyxl", "et_xmlfile", "tkinterdnd2", "lxml", "pillow",
        "pyinstaller", "pyinstaller-hooks-contrib", "altgraph", "packaging",
        "pefile", "pywin32-ctypes", "setuptools",
    ]
    required = {"openpyxl", "et_xmlfile", "tkinterdnd2", "pyinstaller"}
    for name in packages:
        try:
            distribution = metadata.distribution(name)
        except metadata.PackageNotFoundError:
            if name in required:
                raise
            continue
        inventory.append(f"{name}=={distribution.version}")
        count = 0
        for item in distribution.files or []:
            if not any(word in item.name.lower() for word in (
                "license", "licence", "copying", "copyright", "notice"
            )):
                continue
            source = Path(distribution.locate_file(item))
            if not source.is_file():
                continue
            relative = Path(*[p for p in item.parts if p not in ("..", ".")])
            target = notices / name / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
            count += 1
        if count == 0:
            raise RuntimeError(f"No license file found for {name}")

    shutil.copyfile(Path(sys.base_prefix) / "LICENSE.txt", notices / "Python-LICENSE.txt")
    tk_licenses = list((Path(sys.base_prefix) / "tcl").glob("*/license*"))
    if not tk_licenses:
        raise RuntimeError("Tk license file not found.")
    for source in tk_licenses:
        if source.is_file():
            shutil.copyfile(source, notices / f"{source.parent.name}-{source.name}")
    tcl_version = tkinter.Tcl().eval("info patchlevel")
    tag = "core-" + tcl_version.replace(".", "-")
    url = f"https://raw.githubusercontent.com/tcltk/tcl/{tag}/license.terms"
    with urllib.request.urlopen(url, timeout=30) as response:
        license_bytes = response.read()
    if len(license_bytes) < 500 or b"copyright" not in license_bytes.lower():
        raise RuntimeError("Unexpected Tcl license response.")
    (notices / "Tcl-license.terms").write_bytes(license_bytes)
    inventory.append(f"Tcl {tcl_version}; license source: {url}")
    (bundle / "DEPENDENCIES.txt").write_text("\n".join(inventory) + "\n", encoding="utf-8")


def main():
    if sys.platform != "win32":
        raise SystemExit("Package the Windows build on Windows.")
    output = ROOT / "release"
    bundle = output / f"ExcelToiPhoneVCF-{APP_VERSION}-windows-x64"
    bundle.mkdir(exist_ok=True)
    shutil.copyfile(output / "Excel转通讯录.exe", bundle / "ExcelToiPhoneVCF.exe")
    for name in ("LICENSE", "README.md"):
        shutil.copyfile(ROOT / name, bundle / name)
    examples = bundle / "examples"
    examples.mkdir(exist_ok=True)
    shutil.copyfile(ROOT / "examples" / "contacts.csv", examples / "contacts.csv")
    copy_notices(bundle)
    zip_path = output / f"{bundle.name}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(bundle.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(bundle).as_posix())
    digest = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    (output / "SHA256SUMS.txt").write_text(f"{digest}  {zip_path.name}\n", encoding="ascii")
    print(f"Packaged: {zip_path.name}")


if __name__ == "__main__":
    main()
