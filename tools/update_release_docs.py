"""Update a release ZIP's README while preserving its executable and other files."""

import hashlib
from pathlib import Path
import sys
import zipfile


def main():
    archive_path = Path(sys.argv[1]).resolve()
    readme = Path(__file__).resolve().parents[1] / "README.md"
    temporary = archive_path.with_suffix(".updated.zip")
    with zipfile.ZipFile(archive_path) as source:
        if len(source.namelist()) != len(set(source.namelist())):
            raise ValueError("Duplicate ZIP entries.")
        before = {name: hashlib.sha256(source.read(name)).digest() for name in source.namelist()}
        if "README.md" not in before or "ExcelToiPhoneVCF.exe" not in before:
            raise ValueError("Unexpected release package.")
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                payload = readme.read_bytes() if item.filename == "README.md" else source.read(item.filename)
                target.writestr(item, payload)
    with zipfile.ZipFile(temporary) as updated:
        for name, digest in before.items():
            if name != "README.md" and hashlib.sha256(updated.read(name)).digest() != digest:
                raise ValueError(f"Unexpected modification: {name}")
        if updated.read("README.md") != readme.read_bytes():
            raise ValueError("README update failed.")
    temporary.replace(archive_path)
    checksum = hashlib.sha256(archive_path.read_bytes()).hexdigest()
    (archive_path.parent / "SHA256SUMS.txt").write_text(
        f"{checksum}  {archive_path.name}\n", encoding="ascii"
    )
    print("Updated README; executable and all other files unchanged.")


if __name__ == "__main__":
    main()
