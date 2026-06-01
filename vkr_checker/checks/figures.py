"""Проверка оформления рисунков и диаграмм."""
from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .base import BaseCheck, CheckResult, Severity, add_issue


class FiguresCheck(BaseCheck):
    check_id = "figures"
    check_name = "Оформление рисунков"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_captions = self.rules["captions"]
        figure_re = re.compile(rules_captions["figure_pattern"])
        figure_numbers_seen = []
        figure_number_re = re.compile(r"^Рис\.\s+(\d+)\.")

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()
            if not text.startswith("Рис."):
                continue

            # Извлекаем номер
            m = figure_number_re.match(text)
            if m:
                figure_numbers_seen.append(int(m.group(1)))

            # Проверяем формат подписи
            if not figure_re.match(text):
                add_issue(
                    result,
                    rule_id="figure_caption_format",
                    message=(
                        "Неверный формат подписи рисунка. "
                        "Должно быть: «Рис. N. Название» "
                        "(с прописной буквы, без точки в конце)"
                    ),
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                    context=text[:80],
                )
            else:
                # Точка в конце названия недопустима
                if text.endswith("."):
                    add_issue(
                        result,
                        rule_id="figure_caption_dot",
                        message="Подпись рисунка не должна заканчиваться точкой",
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )

            # Выравнивание
            if para.alignment not in (WD_ALIGN_PARAGRAPH.CENTER, None):
                add_issue(
                    result,
                    rule_id="figure_caption_alignment",
                    message="Подпись рисунка должна быть выровнена по центру",
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                )

        # Сквозная нумерация
        for j, num in enumerate(figure_numbers_seen):
            if num != j + 1:
                add_issue(
                    result,
                    rule_id="figure_numbering",
                    message=(
                        f"Нарушена сквозная нумерация рисунков: "
                        f"ожидался Рис. {j+1}, найден Рис. {num}"
                    ),
                    severity=Severity.ERROR,
                )
                break
