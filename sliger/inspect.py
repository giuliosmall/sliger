"""Static inspection of Jinja formulas in a presentation."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from jinja2 import Environment, meta

from sliger.quotes import find_smart_quotes
from sliger.templates import formula_from_element, looks_like_formula

_FUNC_CALL = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*\(")


@dataclass
class FormulaInfo:
    slide_number: int
    object_id: str
    formula: str
    functions: tuple[str, ...]
    variables: tuple[str, ...]
    smart_quotes: tuple[str, ...]


@dataclass
class InspectReport:
    formulas: tuple[FormulaInfo, ...]
    functions: tuple[str, ...]
    variables: tuple[str, ...]
    smart_quotes: bool
    missing_data: tuple[str, ...] = field(default_factory=tuple)

    def to_dict(self) -> dict[str, Any]:
        return {
            "formulas": [
                {
                    "slide_number": item.slide_number,
                    "object_id": item.object_id,
                    "formula": item.formula,
                    "functions": list(item.functions),
                    "variables": list(item.variables),
                    "smart_quotes": list(item.smart_quotes),
                }
                for item in self.formulas
            ],
            "functions": list(self.functions),
            "variables": list(self.variables),
            "smart_quotes": self.smart_quotes,
            "missing_data": list(self.missing_data),
        }


def analyze_formula(formula: str, env: Environment) -> tuple[tuple[str, ...], tuple[str, ...]]:
    ast = env.parse(formula)
    variables = tuple(sorted(meta.find_undeclared_variables(ast)))
    functions = tuple(sorted({match.group(1) for match in _FUNC_CALL.finditer(formula)}))
    return functions, variables


def inspect_slides(
    slides: list[dict[str, Any]],
    env: Environment,
    data: dict[str, Any] | None = None,
    get_text=None,
) -> InspectReport:
    from sliger.slides_utils import get_text_elements_from_slide, gslides_element_to_text

    text_fn = get_text or gslides_element_to_text
    formulas: list[FormulaInfo] = []
    all_functions: set[str] = set()
    all_variables: set[str] = set()
    any_smart = False
    for index, slide in enumerate(slides, start=1):
        page_id = slide.get("objectId") or ""
        for element in get_text_elements_from_slide(slide):
            parsed = text_fn(element, page_id)
            formula = formula_from_element(element, parsed["text"])
            if not looks_like_formula(formula):
                continue
            functions, variables = analyze_formula(formula, env)
            smart = tuple(find_smart_quotes(formula))
            any_smart = any_smart or bool(smart)
            all_functions.update(functions)
            all_variables.update(variables)
            formulas.append(
                FormulaInfo(
                    slide_number=index,
                    object_id=parsed["object_id"],
                    formula=formula,
                    functions=functions,
                    variables=variables,
                    smart_quotes=smart,
                )
            )
    reserved = {
        "now",
        "strftime",
        "sql",
        "sliger_repeat",
        "range",
        "dict",
        "lipsum",
        "cycler",
        "joiner",
        "namespace",
    }
    required = tuple(
        sorted(var for var in all_variables if var not in reserved and var not in all_functions)
    )
    missing = tuple(name for name in required if data is not None and name not in data)
    return InspectReport(
        formulas=tuple(formulas),
        functions=tuple(sorted(all_functions)),
        variables=required,
        smart_quotes=any_smart,
        missing_data=missing,
    )
