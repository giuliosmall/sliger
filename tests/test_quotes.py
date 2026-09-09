from __future__ import annotations

from sliger.quotes import find_smart_quotes, has_smart_quotes, replace_smart_quotes

LEFT_SINGLE = "\u2018"
RIGHT_SINGLE = "\u2019"
LEFT_DOUBLE = "\u201c"
RIGHT_DOUBLE = "\u201d"


def test_find_smart_quotes_collects_unique_sorted() -> None:
    text = f"{LEFT_DOUBLE}hello{RIGHT_DOUBLE} {LEFT_SINGLE}and{RIGHT_SINGLE}"
    assert find_smart_quotes(text) == [
        LEFT_SINGLE,
        RIGHT_SINGLE,
        LEFT_DOUBLE,
        RIGHT_DOUBLE,
    ]


def test_replace_smart_quotes_to_ascii() -> None:
    text = f"{LEFT_SINGLE}hello{RIGHT_SINGLE} {LEFT_DOUBLE}world{RIGHT_DOUBLE}"
    assert replace_smart_quotes(text) == "'hello' \"world\""


def test_has_smart_quotes_true_for_curly() -> None:
    assert has_smart_quotes(f"it{RIGHT_SINGLE}s") is True
    assert has_smart_quotes(f"{LEFT_DOUBLE}quoted{RIGHT_DOUBLE}") is True


def test_clean_ascii_has_no_smart_quotes() -> None:
    text = 'it\'s a "plain" string'
    assert find_smart_quotes(text) == []
    assert has_smart_quotes(text) is False
    assert replace_smart_quotes(text) == text
