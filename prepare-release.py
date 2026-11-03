"""Prepare a redistributable ZIP without local installation records or paths."""
import hashlib
import json
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VERSION = "v1.0.0-zh.1"
DIST = ROOT / "build" / "release"
DIST.mkdir(parents=True, exist_ok=True)

files = [
    "addin/ConvertCsvToTableZh.xlam",
    "addin/Install.vbs",
    "addin/Uninstall.vbs",
    "README.md",
    "README.zh-CN.md",
    "LICENSE",
    "docs/verification.json",
]
for name in files:
    if not (ROOT / name).is_file():
        raise FileNotFoundError(name)

archive_path = DIST / f"ConvertCsvToTable-zh-{VERSION}.zip"
with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
    for name in files:
        archive.write(ROOT / name, f"ConvertCsvToTable-zh-{VERSION}/{name}")

asset_paths = [archive_path, ROOT / "addin" / "ConvertCsvToTableZh.xlam"]
sums_path = DIST / "SHA256SUMS.txt"
sums_path.write_text("".join(f"{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.name}\n" for p in asset_paths), encoding="ascii")
asset_paths.append(sums_path)
print(json.dumps([{"name": p.name, "size": p.stat().st_size, "sha256": hashlib.sha256(p.read_bytes()).hexdigest()} for p in asset_paths], indent=2))
