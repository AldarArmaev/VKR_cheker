"""Проверка наличия обязательных разделов."""
from __future__ import annotations

import re
from docx.oxml.ns import qn
from .base import BaseCheck, CheckResult, Severity, add_issue


def is_numbered_paragraph(para) -> bool:
    """Проверяет, есть ли у параграфа автоматическая нумерация."""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    numPr = pPr.find(qn("w:numPr"))
    return numPr is not None


def get_list_level(para):
    """Возвращает уровень нумерации (ilvl) или None."""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    ilvl = numPr.find(qn("w:ilvl"))
    if ilvl is not None and ilvl.get(qn("w:val")):
        return int(ilvl.get(qn("w:val")))
    return None


class SectionsCheck(BaseCheck):
    """Проверка наличия и порядка обязательных разделов ВКР."""

    check_id = "sections"
    check_name = "Обязательные разделы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        required = self.rules["required_sections"]

        # Сначала ищем разделы стандартным способом (по тексту)
        found_names = {s.name for s in model.sections_found}

        for sec in required:
            pattern = sec["pattern"]
            name = sec["name"]
            # Стандартная проверка по тексту
            found = any(
                re.match(pattern, p.text.strip(), re.IGNORECASE)
                for p in model.paragraphs if p.text.strip()
            )
            # Дополнительная проверка для глав с автоматической нумерацией
            if not found and name.startswith("Глава"):
                for para in model.paragraphs:
                    if is_numbered_paragraph(para) and get_list_level(para) == 0:
                        # Считаем, что это глава (можно дополнительно проверить, что текст не пустой)
                        # Для точности можно также проверить, что стиль параграфа содержит "Заголовок"
                        if para.text.strip():
                            found = True
                            break
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