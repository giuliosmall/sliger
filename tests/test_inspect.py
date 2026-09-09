from __future__ import annotations

from jinja2 import Environment

from sliger.inspect import analyze_formula, inspect_slides
from sliger.templates import FORMULA_PREFIX


def test_analyze_formula_finds_functions_and_variables() -> None:
    functions, variables = analyze_formula("{{ greet(name) }}", Environment())
    assert "greet" in functions
    assert "name" in variables
    assert "greet" in variables


def test_inspect_slides_finds_formula(text_element_factory) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [
            text_element_factory("shape-1", "Hello {{ name }}"),
            text_element_factory("shape-2", "plain text"),
        ],
    }
    report = inspect_slides([slide], Environment(), data={"name": "Ada"})
    assert len(report.formulas) == 1
    info = report.formulas[0]
    assert info.slide_number == 1
    assert info.object_id == "shape-1"
    assert info.formula == "Hello {{ name }}"
    assert info.variables == ("name",)
    assert report.variables == ("name",)
    assert report.missing_data == ()
    assert report.smart_quotes is False


def test_inspect_missing_data_keys(text_element_factory) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "{{ greet(name) }}")],
    }
    report = inspect_slides([slide], Environment(), data={"other": 1})
    assert report.functions == ("greet",)
    assert report.variables == ("name",)
    assert report.missing_data == ("name",)

    complete = inspect_slides([slide], Environment(), data={"name": "Ada"})
    assert complete.missing_data == ()

    no_data = inspect_slides([slide], Environment())
    assert no_data.missing_data == ()


def test_inspect_smart_quotes_flag(text_element_factory) -> None:
    curly = "Hello \u201c{{ name }}\u201d"
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", curly)],
    }
    report = inspect_slides([slide], Environment(), data={"name": "Ada"})
    assert report.smart_quotes is True
    assert "\u201c" in report.formulas[0].smart_quotes
    assert "\u201d" in report.formulas[0].smart_quotes


def test_inspect_stored_formula_in_description(text_element_factory) -> None:
    element = text_element_factory("shape-1", "Visible title")
    element["description"] = FORMULA_PREFIX + "{{ stored_key }}"
    slide = {"objectId": "slide-1", "pageElements": [element]}
    report = inspect_slides([slide], Environment(), data={})
    assert report.formulas[0].formula == "{{ stored_key }}"
    assert report.variables == ("stored_key",)
    assert report.missing_data == ("stored_key",)
    as_dict = report.to_dict()
    assert as_dict["formulas"][0]["object_id"] == "shape-1"
    assert as_dict["missing_data"] == ["stored_key"]
