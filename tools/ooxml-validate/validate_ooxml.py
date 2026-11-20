from __future__ import annotations

import argparse
import copy
import functools
import io
import os
import posixpath
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

from fetch_xsds import ensure_xsds

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

COMMENTS_ORDER = [
    "authors",
    "commentList",
    "extLst",
]

METADATA_ORDER = [
    "metadataTypes",
    "metadataStrings",
    "mdxMetadata",
    "futureMetadata",
    "cellMetadata",
    "valueMetadata",
    "extLst",
]

PIVOT_TABLE_ORDER = [
    "location",
    "pivotFields",
    "rowFields",
    "rowItems",
    "colFields",
    "colItems",
    "pageFields",
    "dataFields",
    "formats",
    "conditionalFormats",
    "chartFormats",
    "pivotHierarchies",
    "pivotTableStyleInfo",
    "filters",
    "rowHierarchiesUsage",
    "colHierarchiesUsage",
    "extLst",
]

ROOT_SEQUENCE_MAP = {
    "workbook": WORKBOOK_ORDER,
    "worksheet": WORKSHEET_ORDER,
    "styleSheet": STYLESHEET_ORDER,
    "table": TABLE_ORDER,
    "comments": COMMENTS_ORDER,
    "metadata": METADATA_ORDER,
    "pivotTableDefinition": PIVOT_TABLE_ORDER,
}

ROOT_XSD_MAP = {
    "workbook": "sml.xsd",
    "worksheet": "sml.xsd",
    "styleSheet": "sml.xsd",
    "sst": "sml.xsd",
    "table": "sml.xsd",
    "comments": "sml.xsd",
    "metadata": "sml.xsd",
    "calcChain": "sml.xsd",
    "pivotTableDefinition": "sml.xsd",
    "pivotCacheDefinition": "sml.xsd",
    "pivotCacheRecords": "sml.xsd",
    "wsDr": "dml-spreadsheetDrawing.xsd",
    "chartSpace": "dml-chart.xsd",
    "Types": "opc-contentTypes.xsd",
    "Relationships": "opc-relationships.xsd",
}

VALID_ERROR_LITERALS = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
}

CELL_REF_RE = re.compile(r"^([A-Za-z]+)([0-9]+)$")
EXTERNAL_SCHEME_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.-]*:")
FORBIDDEN_SHEET_CHARS = set("\\/?*[]:")
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def normalize_part(name: str) -> str:
    return posixpath.normpath(name.replace("\\", "/").lstrip("/"))


def col_letters_to_idx(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def parse_cell_ref(ref: str) -> tuple[int, int] | None:
    m = CELL_REF_RE.match(ref)
    if not m:
        return None
    col = col_letters_to_idx(m.group(1))
    row = int(m.group(2))
    return (row, col)


def parse_range_ref(ref: str) -> tuple[int, int, int, int] | None:
    parts = ref.split(":")
    if len(parts) == 1:
        p = parse_cell_ref(parts[0])
        if p is None:
            return None
        return (p[0], p[1], p[0], p[1])
    if len(parts) == 2:
        p1 = parse_cell_ref(parts[0])
        p2 = parse_cell_ref(parts[1])
        if p1 is None or p2 is None:
            return None
        return (p1[0], p1[1], p2[0], p2[1])
    return None


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


def check_count_attr(elem: etree._Element, part_name: str) -> list[str]:
    count_str = elem.attrib.get("count")
    if count_str is None:
        return []
    try:
        declared = int(count_str)
    except ValueError:
        return [f"{part_name}: <{local_name(elem.tag)}> has non-integer count={count_str!r}"]
    actual = sum(1 for ch in elem if isinstance(ch.tag, str))
    if declared != actual:
        return [
            f"{part_name}: <{local_name(elem.tag)}> count={declared} does not match child count {actual}"
        ]
    return []


XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


@functools.lru_cache(maxsize=8)
def load_xsd_cache(xsd_dir: Path | None) -> dict[str, etree.XMLSchema]:
    cache: dict[str, etree.XMLSchema] = {}
    if xsd_dir is None or not xsd_dir.is_dir():
        return cache

    class LocalXsdResolver(etree.Resolver):
        def resolve(self, url, id, context):
            candidate = xsd_dir / Path(url).name
            if candidate.is_file():
                return self.resolve_filename(str(candidate), context)
            return None

    parser = etree.XMLParser()
    parser.resolvers.add(LocalXsdResolver())

    for xsd_name in sorted(set(ROOT_XSD_MAP.values())):
        xsd_path = xsd_dir / xsd_name
        if xsd_path.is_file():
            doc = etree.parse(str(xsd_path), parser)
            cache[xsd_name] = etree.XMLSchema(doc)
    return cache


def prepare_for_xsd_validation(root: etree._Element, xsd_name: str) -> etree._Element:
    if xsd_name != "sml.xsd":
        return root
    cloned = copy.deepcopy(root)
    for elem in cloned.iter():
        if isinstance(elem.tag, str) and local_name(elem.tag) == "t":
            if elem.attrib.get(XML_SPACE_ATTR) == "preserve":
                del elem.attrib[XML_SPACE_ATTR]
    return cloned


def validate_ooxml_bytes(
    data: bytes,
    xsd_dir: Path | None = None,
    require_xsd: bool = False,
) -> list[str]:
    errors: list[str] = []
    xsd_cache = load_xsd_cache(xsd_dir)
    if require_xsd and not xsd_cache:
        return [f"No XSD schemas loaded from {xsd_dir}"]

    try:
        zf = zipfile.ZipFile(io.BytesIO(data), "r")
    except zipfile.BadZipFile as exc:
        return [f"Invalid ZIP container: {exc}"]

    with zf:
        part_map: dict[str, str] = {}
        for info in zf.infolist():
            if info.is_dir():
                continue
            raw_name = info.filename
            if raw_name.startswith("/"):
                errors.append(f"{raw_name}: OPC part name must not start with '/'")
            if ".." in raw_name.replace("\\", "/").split("/"):
                errors.append(f"{raw_name}: OPC part name must not contain '..'")
            norm_lower = normalize_part(raw_name).lower()
            if norm_lower in part_map:
                errors.append(
                    f"{raw_name}: duplicate OPC part name (case-insensitive collision with {part_map[norm_lower]})"
                )
            else:
                part_map[norm_lower] = raw_name

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
                        ext = child.attrib["Extension"].lower()
                        if ext in default_exts:
                            errors.append(
                                f"[Content_Types].xml: duplicate Default Extension={ext!r}"
                            )
                        default_exts.add(ext)
                    elif lname == "Override" and "PartName" in child.attrib:
                        pname = normalize_part(child.attrib["PartName"]).lower()
                        if pname in override_parts:
                            errors.append(
                                f"[Content_Types].xml: duplicate Override PartName={child.attrib['PartName']!r}"
                            )
                        override_parts.add(pname)
            except etree.XMLSyntaxError as exc:
                errors.append(f"[Content_Types].xml: XML syntax error: {exc}")

        parsed_parts: dict[str, etree._Element] = {}
        rels_by_source: dict[str, dict[str, tuple[str, str, str]]] = {}

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

            parsed_parts[norm_lower] = root
            root_lname = local_name(root.tag)
            if root_lname in ROOT_SEQUENCE_MAP:
                errors.extend(
                    check_sequence(root, orig_name, ROOT_SEQUENCE_MAP[root_lname])
                )

            if root_lname in ROOT_XSD_MAP:
                xsd_name = ROOT_XSD_MAP[root_lname]
                schema = xsd_cache.get(xsd_name)
                if schema is not None:
                    target_root = prepare_for_xsd_validation(root, xsd_name)
                    if not schema.validate(target_root):
                        for log_entry in schema.error_log:
                            errors.append(
                                f"{orig_name} ({xsd_name}:{log_entry.line}): {log_entry.message}"
                            )

            if norm_lower.endswith(".rels"):
                src = source_part_for_rels(orig_name).lower()
                seen_ids: dict[str, tuple[str, str, str]] = {}
                for rel in root:
                    if not isinstance(rel.tag, str):
                        continue
                    if local_name(rel.tag) != "Relationship":
                        continue
                    rid = rel.attrib.get("Id", "")
                    rtype = rel.attrib.get("Type", "")
                    target = rel.attrib.get("Target", "")
                    mode = rel.attrib.get("TargetMode", "")
                    if not rid:
                        errors.append(f"{orig_name}: Relationship missing Id")
                    elif rid in seen_ids:
                        errors.append(f"{orig_name}: duplicate Relationship Id={rid!r}")
                    else:
                        seen_ids[rid] = (rtype, target, mode)
                    if not rtype:
                        errors.append(f"{orig_name}: Relationship {rid!r} missing Type")
                    if mode.lower() == "external":
                        continue
                    if EXTERNAL_SCHEME_RE.match(target):
                        errors.append(
                            f"{orig_name}: Relationship {rid!r} has URI target {target!r} without TargetMode='External'"
                        )
                        continue
                    resolved = resolve_target(src, target).lower()
                    if resolved not in part_map:
                        errors.append(
                            f"{orig_name}: dangling relationship {rid} -> {target} ({resolved})"
                        )
                rels_by_source[src] = seen_ids

        sst_count: int | None = None
        for norm_lower, root in parsed_parts.items():
            if local_name(root.tag) == "sst":
                sst_count = sum(
                    1 for ch in root if isinstance(ch.tag, str) and local_name(ch.tag) == "si"
                )

        cell_xfs_count: int | None = None
        for norm_lower, root in parsed_parts.items():
            if local_name(root.tag) != "styleSheet":
                continue
            orig_name = part_map[norm_lower]
            for child in root:
                if not isinstance(child.tag, str):
                    continue
                lname = local_name(child.tag)
                if lname in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs"):
                    errors.extend(check_count_attr(child, orig_name))
                if lname == "fills":
                    fill_count = sum(1 for ch in child if isinstance(ch.tag, str))
                    if fill_count < 2:
                        errors.append(
                            f"{orig_name}: <fills> must contain at least 2 entries (none and gray125), found {fill_count}"
                        )
                if lname == "cellXfs":
                    cell_xfs_count = sum(1 for ch in child if isinstance(ch.tag, str))
                    if cell_xfs_count < 1:
                        errors.append(f"{orig_name}: <cellXfs> must contain at least 1 <xf>")

        for norm_lower, root in parsed_parts.items():
            if local_name(root.tag) != "workbook":
                continue
            orig_name = part_map[norm_lower]
            wb_rels = rels_by_source.get(norm_lower, {})
            seen_names: set[str] = set()
            seen_ids: set[str] = set()
            for child in root:
                if not isinstance(child.tag, str) or local_name(child.tag) != "sheets":
                    continue
                for sheet_el in child:
                    if not isinstance(sheet_el.tag, str) or local_name(sheet_el.tag) != "sheet":
                        continue
                    sname = sheet_el.attrib.get("name", "")
                    sid = sheet_el.attrib.get("sheetId", "")
                    rid = (
                        sheet_el.attrib.get(f"{{{REL_NS}}}id")
                        or sheet_el.attrib.get(f"{{{REL_NS_STRICT}}}id")
                        or ""
                    )
                    if not sname or len(sname) > 31 or any(c in FORBIDDEN_SHEET_CHARS for c in sname):
                        errors.append(f"{orig_name}: invalid sheet name {sname!r}")
                    elif sname.lower() in seen_names:
                        errors.append(f"{orig_name}: duplicate sheet name {sname!r}")
                    seen_names.add(sname.lower())
                    if not sid.isdigit() or int(sid) < 1:
                        errors.append(f"{orig_name}: invalid sheetId {sid!r} for {sname!r}")
                    elif sid in seen_ids:
                        errors.append(f"{orig_name}: duplicate sheetId {sid!r}")
                    seen_ids.add(sid)
                    if rid and rid not in wb_rels:
                        errors.append(
                            f"{orig_name}: sheet {sname!r} references missing relationship {rid!r}"
                        )

        for norm_lower, root in parsed_parts.items():
            if local_name(root.tag) != "worksheet":
                continue
            orig_name = part_map[norm_lower]
            for child in root:
                if not isinstance(child.tag, str):
                    continue
                lname = local_name(child.tag)
                if lname == "mergeCells":
                    errors.extend(check_count_attr(child, orig_name))
                    rects: list[tuple[int, int, int, int, str]] = []
                    for mc in child:
                        if not isinstance(mc.tag, str) or local_name(mc.tag) != "mergeCell":
                            continue
                        mref = mc.attrib.get("ref", "")
                        parsed = parse_range_ref(mref)
                        if parsed is None:
                            errors.append(f"{orig_name}: invalid mergeCell ref {mref!r}")
                            continue
                        r1, c1, r2, c2 = parsed
                        for pr1, pc1, pr2, pc2, pref in rects:
                            if not (r2 < pr1 or r1 > pr2 or c2 < pc1 or c1 > pc2):
                                errors.append(
                                    f"{orig_name}: overlapping mergeCells {pref!r} and {mref!r}"
                                )
                        rects.append((r1, c1, r2, c2, mref))
                elif lname == "sheetData":
                    prev_row = 0
                    for row_el in child:
                        if not isinstance(row_el.tag, str) or local_name(row_el.tag) != "row":
                            continue
                        r_attr = row_el.attrib.get("r")
                        cur_row = prev_row + 1
                        if r_attr is not None:
                            if not r_attr.isdigit() or int(r_attr) < 1:
                                errors.append(f"{orig_name}: invalid row/@r={r_attr!r}")
                                continue
                            cur_row = int(r_attr)
                            if cur_row <= prev_row:
                                errors.append(
                                    f"{orig_name}: row/@r={cur_row} not strictly ascending after {prev_row}"
                                )
                        prev_row = cur_row
                        prev_col = 0
                        for c_el in row_el:
                            if not isinstance(c_el.tag, str) or local_name(c_el.tag) != "c":
                                continue
                            cref = c_el.attrib.get("r")
                            if cref is not None:
                                parsed = parse_cell_ref(cref)
                                if parsed is None:
                                    errors.append(f"{orig_name}: invalid cell ref {cref!r}")
                                else:
                                    crow, ccol = parsed
                                    if crow != cur_row:
                                        errors.append(
                                            f"{orig_name}: cell {cref!r} row {crow} != parent row {cur_row}"
                                        )
                                    if ccol <= prev_col:
                                        errors.append(
                                            f"{orig_name}: cell {cref!r} not strictly ascending in row {cur_row}"
                                        )
                                    prev_col = ccol
                            s_attr = c_el.attrib.get("s")
                            if s_attr is not None and cell_xfs_count is not None:
                                if not s_attr.isdigit() or int(s_attr) >= cell_xfs_count:
                                    errors.append(
                                        f"{orig_name}: cell {cref or '?'} s={s_attr!r} out of bounds (cellXfs={cell_xfs_count})"
                                    )
                            t_attr = c_el.attrib.get("t", "n")
                            v_text = None
                            for sub in c_el:
                                if isinstance(sub.tag, str) and local_name(sub.tag) == "v":
                                    v_text = sub.text or ""
                                    break
                            if t_attr == "s" and v_text is not None and sst_count is not None:
                                if not v_text.isdigit() or int(v_text) >= sst_count:
                                    errors.append(
                                        f"{orig_name}: cell {cref or '?'} shared string index {v_text!r} out of bounds (sst={sst_count})"
                                    )
                            elif t_attr == "b" and v_text is not None:
                                if v_text not in ("0", "1"):
                                    errors.append(
                                        f"{orig_name}: cell {cref or '?'} boolean value {v_text!r} must be '0' or '1'"
                                    )
                            elif t_attr == "e" and v_text is not None:
                                if v_text not in VALID_ERROR_LITERALS:
                                    errors.append(
                                        f"{orig_name}: cell {cref or '?'} error value {v_text!r} is not a valid ECMA-376 error literal"
                                    )

        for norm_lower, root in parsed_parts.items():
            if local_name(root.tag) != "comments":
                continue
            orig_name = part_map[norm_lower]
            seen_refs: set[str] = set()
            for child in root:
                if not isinstance(child.tag, str) or local_name(child.tag) != "commentList":
                    continue
                for cm in child:
                    if not isinstance(cm.tag, str) or local_name(cm.tag) != "comment":
                        continue
                    cref = cm.attrib.get("ref", "").upper()
                    if not cref or parse_cell_ref(cref) is None:
                        errors.append(f"{orig_name}: invalid comment ref {cref!r}")
                    elif cref in seen_refs:
                        errors.append(f"{orig_name}: duplicate comment ref {cref!r}")
                    seen_refs.add(cref)

    return errors


def validate_ooxml_file(
    path: Path,
    xsd_dir: Path | None = None,
    require_xsd: bool = False,
) -> list[str]:
    return validate_ooxml_bytes(
        path.read_bytes(), xsd_dir=xsd_dir, require_xsd=require_xsd
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate .xlsx/.xlsm files against ECMA-376 OPC rules, semantic invariants, and XSDs."
    )
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument(
        "--xsd-dir",
        type=Path,
        default=Path(os.environ["OOXML_XSD_DIR"])
        if os.environ.get("OOXML_XSD_DIR")
        else None,
    )
    parser.add_argument("--no-xsd", action="store_true")
    parser.add_argument("--require-xsd", action="store_true")
    args = parser.parse_args()

    xsd_dir = None if args.no_xsd else ensure_xsds(args.xsd_dir)
    require_xsd = not args.no_xsd or args.require_xsd

    failed = False
    for file_path in args.files:
        errs = validate_ooxml_file(
            file_path, xsd_dir=xsd_dir, require_xsd=require_xsd
        )
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
