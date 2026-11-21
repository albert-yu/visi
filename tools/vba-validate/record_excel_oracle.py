from __future__ import annotations

import json
import sys
from pathlib import Path


def record_with_excel(golden_path: Path) -> None:
    try:
        import xlwings as xw
    except ImportError:
        sys.exit("xlwings is required to record golden outputs against Microsoft Excel.")

    cases = json.loads(golden_path.read_text(encoding="utf-8"))
    app = xw.App(visible=False)
    try:
        for case in cases:
            wb = app.books.add()
            try:
                mod = wb.api.VBProject.VBComponents.Add(1)
                mod.CodeModule.AddFromString(case["source"])
                wrapper = (
                    "Function __OracleWrap()\n"
                    f"    Dim r: r = {case['procedure']}()\n"
                    "    __OracleWrap = TypeName(r) & \"|\" & CStr(r)\n"
                    "End Function\n"
                )
                mod.CodeModule.AddFromString(wrapper)
                raw = wb.macro("__OracleWrap")()
                tname, val = raw.split("|", 1)
                case["expected_type"] = tname
                case["expected_value"] = val
            finally:
                wb.close()
    finally:
        app.quit()

    golden_path.write_text(json.dumps(cases, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "golden_cases.json"
    record_with_excel(target)
