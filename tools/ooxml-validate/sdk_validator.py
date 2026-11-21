from __future__ import annotations

import argparse
import functools
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Sequence

SDK_PROJECT_DIR = Path(__file__).resolve().parent / "sdk-validator"
SDK_CSPROJ = SDK_PROJECT_DIR / "OpenXmlSdkValidator.csproj"
SDK_PROGRAM = SDK_PROJECT_DIR / "Program.cs"
SDK_DLL = SDK_PROJECT_DIR / "bin" / "Release" / "net8.0" / "OpenXmlSdkValidator.dll"


def is_dotnet_available() -> bool:
    return shutil.which("dotnet") is not None


def is_sdk_required() -> bool:
    return os.environ.get("REQUIRE_OPENXML_SDK", "").lower() in ("1", "true", "yes")


@functools.lru_cache(maxsize=1)
def ensure_sdk_validator_built() -> Path:
    if not is_dotnet_available():
        raise RuntimeError("dotnet SDK is not installed or not on PATH")

    needs_build = not SDK_DLL.is_file()
    if not needs_build:
        dll_mtime = SDK_DLL.stat().st_mtime
        for src in (SDK_CSPROJ, SDK_PROGRAM):
            if src.stat().st_mtime > dll_mtime:
                needs_build = True
                break

    if needs_build:
        res = subprocess.run(
            ["dotnet", "build", str(SDK_CSPROJ), "-c", "Release", "--nologo", "-v", "q"],
            cwd=SDK_PROJECT_DIR,
            capture_output=True,
            text=True,
            check=False,
        )
        if res.returncode != 0:
            raise RuntimeError(
                f"dotnet build failed (exit {res.returncode}):\n{res.stdout}\n{res.stderr}"
            )
    return SDK_DLL


def _format_issue(issue: dict[str, str]) -> str:
    part = issue.get("part", "")
    path = issue.get("path", "")
    location = f"{part} {path}".strip() if part and path else (part or path)
    err_type = issue.get("errorType", "")
    err_id = issue.get("id", "")
    desc = issue.get("description", "")
    prefix = f"[{err_type}/{err_id}]" if err_id else f"[{err_type}]"
    if location:
        return f"{prefix} {location}: {desc}"
    return f"{prefix} {desc}"


def validate_files_with_openxml_sdk(
    paths: Sequence[Path | str],
    version: str = "Microsoft365",
) -> dict[str, list[str]]:
    if not paths:
        return {}
    dll = ensure_sdk_validator_built()
    str_paths = [str(p) for p in paths]
    res = subprocess.run(
        ["dotnet", str(dll), "--json", "--version", version, *str_paths],
        capture_output=True,
        text=True,
        check=False,
    )
    if res.returncode not in (0, 1):
        raise RuntimeError(
            f"OpenXmlSdkValidator failed (exit {res.returncode}):\n{res.stdout}\n{res.stderr}"
        )
    report = json.loads(res.stdout)
    out: dict[str, list[str]] = {}
    for entry in report.get("results", []):
        out[entry["file"]] = [_format_issue(issue) for issue in entry.get("errors", [])]
    return out


def validate_file_with_openxml_sdk(
    path: Path | str,
    version: str = "Microsoft365",
) -> list[str]:
    key = str(path)
    results = validate_files_with_openxml_sdk([path], version=version)
    return results.get(key, [])


def validate_bytes_with_openxml_sdk(
    data: bytes,
    suffix: str = ".xlsx",
    version: str = "Microsoft365",
) -> list[str]:
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(data)
        tmp_path = Path(tmp.name)
    try:
        return validate_file_with_openxml_sdk(tmp_path, version=version)
    finally:
        tmp_path.unlink(missing_ok=True)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Validate OOXML workbooks with Microsoft Open XML SDK OpenXmlValidator."
    )
    parser.add_argument("files", nargs="+", help="Paths to .xlsx/.xlsm files to validate")
    parser.add_argument(
        "--version",
        default="Microsoft365",
        help="FileFormatVersions target (default: Microsoft365)",
    )
    args = parser.parse_args(argv)

    results = validate_files_with_openxml_sdk(args.files, version=args.version)
    has_errors = False
    for file_path in args.files:
        errors = results.get(file_path, [])
        if errors:
            has_errors = True
            print(f"FAIL: {file_path} ({len(errors)} error(s))", file=sys.stderr)
            for err in errors:
                print(f"  {err}", file=sys.stderr)
        else:
            print(f"OK: {file_path}")
    return 1 if has_errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
