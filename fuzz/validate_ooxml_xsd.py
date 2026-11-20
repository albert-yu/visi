from __future__ import annotations

import argparse
import posixpath
import sys
import zipfile
from pathlib import Path

from lxml import etree

WORKBOOK_ORDER = [
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
]

WORKSHEET_ORDER = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
]

STYLESHEET_ORDER = [
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
]

TABLE_ORDER = [
    "autoFilter",
    "sortState",
    "tableColumns",
    "tableStyleInfo",
    "extLst",
]

ROOT_SEQUENCE_MAP = {
    "workbook": WORKBOOK_ORDER,
    "worksheet": WORKSHEET_ORDER,
    "styleSheet": STYLESHEET_ORDER,
    "table": TABLE_ORDER,
}

ROOT_XSD_MAP = {
    "workbook": "sml.xsd",
    "worksheet": "sml.xsd",
    "styleSheet": "sml.xsd",
    "sst": "sml.xsd",
    "table": "sml.xsd",
    "pivotTableDefinition": "sml.xsd",
    "pivotCacheDefinition": "sml.xsd",
    "pivotCacheRecords": "sml.xsd",
    "wsDr": "dml-spreadsheetDrawing.xsd",
    "chartSpace": "dml-chart.xsd",
    "Types": "opc-contentTypes.xsd",
    "Relationships": "opc-relationships.xsd",
}


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def normalize_part(name: str) -> str:
    return posixpath.normpath(name.replace("\\", "/").lstrip("/"))


def source_part_for_rels(rels_part: str) -> str:
    norm = normalize_part(rels_part)
    if norm.lower() == "_rels/.rels":
        return ""
    without_ext = norm[:-5] if norm.lower().endswith(".rels") else norm
    marker = "/_rels/"
    if marker in without_ext:
        prefix, rest = without_ext.rsplit(marker, 1)
        return f"{prefix}/{rest}"
    if without_ext.startswith("_rels/"):
        return without_ext[len("_rels/") :]
    return without_ext


def resolve_target(source_part: str, target: str) -> str:
    cleaned = target.replace("\\", "/")
    if cleaned.startswith("/"):
        return normalize_part(cleaned)
    base_dir = posixpath.dirname(normalize_part(source_part)) if source_part else ""
    joined = f"{base_dir}/{cleaned}" if base_dir else cleaned
    return normalize_part(joined)


def check_sequence(root: etree._Element, part_name: str, order: list[str]) -> list[str]:
    errors: list[str] = []
    rank_map = {name: idx for idx, name in enumerate(order)}
    last_rank = -1
    last_name = ""
    for child in root:
        if not isinstance(child.tag, str):
            continue
        lname = local_name(child.tag)
        if lname in rank_map:
            rank = rank_map[lname]
            if rank < last_rank:
                errors.append(
                    f"{part_name}: <{lname}> appears after <{last_name}> (violates XSD sequence)"
                )
            last_rank = rank
            last_name = lname
    return errors


def load_xsd_cache(xsd_dir: Path | None) -> dict[str, etree.XMLSchema]:
    cache: dict[str, etree.XMLSchema] = {}
    if xsd_dir is None or not xsd_dir.is_dir():
        return cache
    for xsd_name in set(ROOT_XSD_MAP.values()):
        xsd_path = xsd_dir / xsd_name
        if xsd_path.is_file():
            doc = etree.parse(str(xsd_path))
            cache[xsd_name] = etree.XMLSchema(doc)
    return cache


def validate_ooxml_bytes(
    data: bytes, xsd_dir: Path | None = None
) -> list[str]:
    import io

    errors: list[str] = []
    xsd_cache = load_xsd_cache(xsd_dir)

    try:
        zf = zipfile.ZipFile(io.BytesIO(data), "r")
    except zipfile.BadZipFile as exc:
        return [f"Invalid ZIP container: {exc}"]

    with zf:
        part_map = {
            normalize_part(info.filename).lower(): info.filename
            for info in zf.infolist()
            if not info.is_dir()
        }

        if "[content_types].xml" not in part_map:
            errors.append("Missing [Content_Types].xml")
        if "_rels/.rels" not in part_map:
            errors.append("Missing _rels/.rels")

        default_exts: set[str] = set()
        override_parts: set[str] = set()

        if "[content_types].xml" in part_map:
            try:
                ct_root = etree.fromstring(zf.read(part_map["[content_types].xml"]))
                for child in ct_root:
                    if not isinstance(child.tag, str):
                        continue
                    lname = local_name(child.tag)
                    if lname == "Default" and "Extension" in child.attrib:
                        default_exts.add(child.attrib["Extension"].lower())
                    elif lname == "Override" and "PartName" in child.attrib:
                        override_parts.add(
                            normalize_part(child.attrib["PartName"]).lower()
                        )
            except etree.XMLSyntaxError as exc:
                errors.append(f"[Content_Types].xml: XML syntax error: {exc}")

        for norm_lower, orig_name in part_map.items():
            if norm_lower != "[content_types].xml":
                ext = norm_lower.rsplit(".", 1)[-1] if "." in norm_lower else ""
                if norm_lower not in override_parts and ext not in default_exts:
                    errors.append(
                        f"{orig_name}: no matching Default or Override in [Content_Types].xml"
                    )

            if not norm_lower.endswith((".xml", ".rels")):
                continue

            raw = zf.read(orig_name)
            try:
                root = etree.fromstring(raw)
            except etree.XMLSyntaxError as exc:
                errors.append(f"{orig_name}: XML syntax error: {exc}")
                continue

            root_lname = local_name(root.tag)
            if root_lname in ROOT_SEQUENCE_MAP:
                errors.extend(
                    check_sequence(root, orig_name, ROOT_SEQUENCE_MAP[root_lname])
                )

            if root_lname in ROOT_XSD_MAP:
                xsd_name = ROOT_XSD_MAP[root_lname]
                schema = xsd_cache.get(xsd_name)
                if schema is not None and not schema.validate(root):
                    for log_entry in schema.error_log:
                        errors.append(
                            f"{orig_name} ({xsd_name}:{log_entry.line}): {log_entry.message}"
                        )

            if norm_lower.endswith(".rels"):
                src = source_part_for_rels(orig_name)
                for rel in root:
                    if not isinstance(rel.tag, str):
                        continue
                    if local_name(rel.tag) != "Relationship":
                        continue
                    mode = rel.attrib.get("TargetMode", "")
                    if mode.lower() == "external":
                        continue
                    target = rel.attrib.get("Target", "")
                    resolved = resolve_target(src, target).lower()
                    if resolved not in part_map:
                        errors.append(
                            f"{orig_name}: dangling relationship {rel.attrib.get('Id', '?')} -> {target} ({resolved})"
                        )

    return errors


def validate_ooxml_file(path: Path, xsd_dir: Path | None = None) -> list[str]:
    return validate_ooxml_bytes(path.read_bytes(), xsd_dir=xsd_dir)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate .xlsx/.xlsm files against OPC rules, XSD element sequences, and optional ECMA-376 XSDs."
    )
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--xsd-dir", type=Path, default=None)
    args = parser.parse_args()

    failed = False
    for file_path in args.files:
        errs = validate_ooxml_file(file_path, xsd_dir=args.xsd_dir)
        if errs:
            failed = True
            print(f"FAIL {file_path}:")
            for err in errs:
                print(f"  - {err}")
        else:
            print(f"OK   {file_path}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
