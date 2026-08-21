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


def test_numeric_looking_text_is_not_equal_to_numbers():
    comparator = DifferentialComparator()

    assert not comparator.values_equal("08", 8)
    assert not comparator.values_equal(8, "08")
    assert not comparator.values_equal("1", 1)
    assert not comparator.values_equal(1, "1")
    assert not comparator.values_equal(".0394", 0.0394)
    assert not comparator.values_equal(0.0394, ".0394")
