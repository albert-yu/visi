from fuzz_excel import DifferentialComparator


def test_text_whitespace_is_significant():
    comparator = DifferentialComparator()

    assert not comparator.values_equal(" a", "a")
    assert not comparator.values_equal("a ", "a")
    assert not comparator.values_equal(" a ", "a")


def test_blank_equivalence_still_allows_missing_whitespace_only_cell():
    comparator = DifferentialComparator()

    assert comparator.values_equal(None, "")
    assert comparator.values_equal(None, "   ")
