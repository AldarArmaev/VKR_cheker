"""Проверка абзацных отступов."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import is_heading_paragraph, para_is_in_table


class IndentsCheck(BaseCheck):
    check_id = "indents"
    check_name = "Абзацные отступы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["indents"]
        body_indent = rules["body_first_line"]
        tol = rules.get("tolerance", 0.1)

        in_main_text = False
        last_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True
            if not in_main_text or not text:
                continue
            if para_is_in_table(para):
                continue

            indent = resolver.get_first_line_indent_cm(para)
            is_heading = is_heading_paragraph(para)

            if is_heading:
                # Заголовки должны быть БЕЗ отступа
                if indent > tol:
                    add_issue(
                        result,
                        rule_id="heading_indent",
                        message=(
                            f"Заголовок имеет красную строку {indent:.2f} см "
                            f"(должно быть 0)"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )
            else:
                # Основной текст — отступ 1.25 см
                if abs(indent - body_indent) > tol:
                    if i - last_issue_para >= 5:
                        add_issue(
                            result,
                            rule_id="body_indent",
                            message=(
                                f"Красная строка {indent:.2f} см "
                                f"(требуется {body_indent} см)"
                            ),
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_issue_para = i
