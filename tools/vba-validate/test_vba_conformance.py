from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

TOOLS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = TOOLS_DIR.parents[1]
GOLDEN_CASES = json.loads((TOOLS_DIR / "golden_cases.json").read_text(encoding="utf-8"))


def _run_via_visi_core(source: str, procedure: str) -> tuple[str, str | None]:
    try:
        import visi_core

        return visi_core.run_macro(source, procedure)
    except ImportError:
        pass

    visi_bin = PROJECT_ROOT / "target" / "debug" / "visi"
    if not visi_bin.exists():
        subprocess.run(
            ["cargo", "build", "-p", "visi"],
            cwd=PROJECT_ROOT,
            check=True,
            capture_output=True,
        )
    res = subprocess.run(
        [str(visi_bin), "macro", "run", "-n", procedure, "--json", "-q", "-"],
        input=source,
        text=True,
        capture_output=True,
        check=True,
    )
    payload = json.loads(res.stdout)
    return payload["type"], payload.get("value")


@pytest.mark.parametrize("case", GOLDEN_CASES, ids=[c["id"] for c in GOLDEN_CASES])
def test_vba_golden_conformance(case: dict) -> None:
    type_name, value = _run_via_visi_core(case["source"], case["procedure"])
    assert type_name == case["expected_type"], f"{case['id']} ({case['spec']}): type mismatch"
    assert value == case["expected_value"], f"{case['id']} ({case['spec']}): value mismatch"
