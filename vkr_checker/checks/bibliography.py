"""Проверка списка литературы."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue


class BibliographyCheck(BaseCheck):
    check_id = "bibliography"
    check_name = "Список литературы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_bib = self.rules["bibliography"]
        min_sources = rules_bib["min_sources"]
        forbidden_types = [ft.lower() for ft in rules_bib["forbidden_source_types"]]

        # Найти раздел "Список литературы"
        bib_start = -1
        for i, para in enumerate(model.paragraphs):
            if re.match(r"^Список литературы$", para.text.strip(), re.IGNORECASE):
                bib_start = i
                break

        if bib_start < 0:
            add_issue(
                result,
                rule_id="bibliography_not_found",
                message="Раздел «Список литературы» не найден",
                severity=Severity.ERROR,
            )
            return

        # Найти конец раздела (следующий основной раздел или конец документа)
        bib_end = len(model.paragraphs)
        for i in range(bib_start + 1, len(model.paragraphs)):
            text = model.paragraphs[i].text.strip()
            if re.match(r"^Приложени", text, re.IGNORECASE):
                bib_end = i
                break

        bib_paras = model.paragraphs[bib_start + 1 : bib_end]

        # Подсчёт источников (непустые строки, начинающиеся с цифры или буквы)
        source_lines = [
            p for p in bib_paras
            if p.text.strip() and not p.text.strip().startswith("#")
        ]
        source_count = len(source_lines)

        if source_count < min_sources:
            add_issue(
                result,
                rule_id="bibliography_count",
                message=(
                    f"В списке литературы {source_count} источников "
                    f"(минимум {min_sources})"
                ),
                severity=Severity.ERROR,
            )

        # Проверка на наличие запрещённых типов источников
        for para in source_lines:
            text_lower = para.text.lower()
            for forbidden in forbidden_types:
                if forbidden in text_lower:
                    add_issue(
                        result,
                        rule_id="bibliography_forbidden_type",
                        message=(
                            f"Запрещённый тип источника: «{forbidden}». "
                            f"Учебники и учебные пособия не включаются в список"
                        ),
                        severity=Severity.ERROR,
                        context=para.text[:80],
                    )
