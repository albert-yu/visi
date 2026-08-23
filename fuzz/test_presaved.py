import glob
import os
import sys
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fuzz_presaved import DEFAULT_PRESAVED_DIR, compare_presaved_file, find_presaved_files
from visi_driver import bindings_available

pytestmark = pytest.mark.skipif(
    not bindings_available(),
    reason="build the bindings with `maturin develop -m visi-python/Cargo.toml --release`",
)

PRESAVED_FILES = find_presaved_files(DEFAULT_PRESAVED_DIR)


@pytest.mark.parametrize(
    "file_path",
    PRESAVED_FILES,
    ids=[os.path.basename(p) for p in PRESAVED_FILES],
)
def test_presaved_file_matches_excel_or_mock(file_path):
    driver = "applescript" if os.path.exists("/Applications/Microsoft Excel.app") else "mock"
    is_match, mismatches, stats = compare_presaved_file(
        file_path=file_path,
        driver_type=driver,
        excel_path="/Applications/Microsoft Excel.app",
        backend="auto",
    )
    assert is_match, f"Found mismatches: {mismatches[:5]}"
    assert stats["total_cells"] > 0
