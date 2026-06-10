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
        figure_pattern_str = rules_captions["figure_pattern"]
        figure_prefix = rules_captions["figure_prefix"]
        figure_re = re.compile(figure_pattern_str)

        # Регулярное выражение для извлечения номера из подписи
        # Ожидается, что номер идёт сразу после префикса и пробела, затем точка
        # Пример: "Рис. 1. Название"
        figure_number_re = re.compile(rf"^{re.escape(figure_prefix)}\s+(\d+)\.")

        figure_numbers_seen = []

        for i, para in enumerate(model.paragraphs):
            # Нормализуем неразрывные пробелы
            text = para.text.strip().replace("\u00A0", " ")
            if not text.startswith(figure_prefix):
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
                        f"Должно быть: «{figure_prefix} N. Название» "
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
                        f"ожидался {figure_prefix} {j+1}, найден {figure_prefix} {num}"
                    ),
                    severity=Severity.ERROR,
                )
                break