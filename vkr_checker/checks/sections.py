"""Проверка наличия обязательных разделов."""
from __future__ import annotations

import re

from .base import BaseCheck, CheckResult, Severity, add_issue


class SectionsCheck(BaseCheck):
    """Проверка наличия и порядка обязательных разделов ВКР."""

    check_id = "sections"
    check_name = "Обязательные разделы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        required = self.rules["required_sections"]
        all_texts = [p.text.strip() for p in model.paragraphs]

        found_names = {s.name for s in model.sections_found}

        for sec in required:
            pattern = sec["pattern"]
            name = sec["name"]
            found = any(
                re.match(pattern, t, re.IGNORECASE)
                for t in all_texts if t
            )
            if not found:
                add_issue(
                    result,
                    rule_id=f"missing_section_{name.replace(' ', '_')}",
                    message=f"Отсутствует обязательный раздел: «{name}»",
                    severity=Severity.ERROR,
                )

        # Проверка порядка разделов
        self._check_order(model, result)

    def _check_order(self, model, result: CheckResult) -> None:
        """Разделы должны идти в предписанном порядке."""
        sections = model.sections_found
        if len(sections) < 2:
            return

        indices = [s.paragraph_index for s in sections]
        for j in range(len(indices) - 1):
            if indices[j] > indices[j + 1]:
                add_issue(
                    result,
                    rule_id="section_order",
                    message=(
                        f"Нарушен порядок разделов: «{sections[j].name}» "
                        f"стоит после «{sections[j+1].name}»"
                    ),
                    severity=Severity.ERROR,
                )
