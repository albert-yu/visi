from __future__ import annotations

import argparse
import io
import os
import sys
import urllib.request
import zipfile
from pathlib import Path

DEFAULT_XSD_DIR = Path(__file__).resolve().parent / ".xsds"

ECMA_ARCHIVES = [
    "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip",
    "https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip",
]

W3C_XML_XSD_URL = "https://www.w3.org/2001/xml.xsd"

REQUIRED_XSDS = (
    "sml.xsd",
    "dml-chart.xsd",
    "dml-spreadsheetDrawing.xsd",
    "opc-contentTypes.xsd",
    "opc-relationships.xsd",
    "xml.xsd",
)


def _extract_xsds_from_zip_bytes(data: bytes, target_dir: Path) -> None:
    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name_lower = info.filename.lower()
            if name_lower.endswith(".xsd"):
                base = Path(info.filename).name
                (target_dir / base).write_bytes(zf.read(info.filename))
            elif name_lower.endswith(".zip") and "xsd" in name_lower or (
                name_lower.endswith(".zip") and "xmlschema" in name_lower
            ):
                _extract_xsds_from_zip_bytes(zf.read(info.filename), target_dir)


def _download_bytes(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "visi-ooxml-validate/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def has_required_xsds(xsd_dir: Path) -> bool:
    return xsd_dir.is_dir() and all((xsd_dir / name).is_file() for name in REQUIRED_XSDS)


def ensure_xsds(xsd_dir: Path | None = None) -> Path:
    dest = xsd_dir or (
        Path(os.environ["OOXML_XSD_DIR"])
        if os.environ.get("OOXML_XSD_DIR")
        else DEFAULT_XSD_DIR
    )
    if has_required_xsds(dest):
        return dest

    dest.mkdir(parents=True, exist_ok=True)

    for url in ECMA_ARCHIVES:
        archive_bytes = _download_bytes(url)
        _extract_xsds_from_zip_bytes(archive_bytes, dest)

    xml_xsd_path = dest / "xml.xsd"
    if not xml_xsd_path.is_file():
        xml_xsd_path.write_bytes(_download_bytes(W3C_XML_XSD_URL))

    missing = [name for name in REQUIRED_XSDS if not (dest / name).is_file()]
    if missing:
        raise RuntimeError(
            f"Missing required ECMA-376 XSD files in {dest} after download: {', '.join(missing)}"
        )

    return dest


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Download ECMA-376 Transitional and OPC XSD schemas locally if missing."
    )
    parser.add_argument("--xsd-dir", type=Path, default=None)
    args = parser.parse_args()
    dest = ensure_xsds(args.xsd_dir)
    print(f"ECMA-376 XSDs ready in {dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
