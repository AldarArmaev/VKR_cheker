"""Проверка межстрочного интервала."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import para_is_in_table


class SpacingCheck(BaseCheck):
    check_id = "spacing"
    check_name = "Межстрочный интервал"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["spacing"]
        body_spacing = rules["body"]
        tol = rules.get("tolerance", 0.05)

        # Проверяем наличие раздела «Введение» — без него проверки невозможны
        if model.intro_start_idx < 0:
            result.skipped = True
            result.skip_reason = (
                "Раздел «Введение» не найден. "
                "Проверка интервалов требует наличия этого раздела."
            )
            return

        in_main_text = False
        last_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            # Начало основного текста — с Введения
            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True
                continue  # Пропускаем сам заголовок "Введение"

            if not in_main_text:
                continue
            if not text:
                continue
            if para_is_in_table(para):
                continue  # таблицы проверяются отдельно

            spacing = resolver.get_line_spacing(para)
            
            # Если интервал не определён (None), просто пропускаем параграф
            # без добавления предупреждения (исправление бага №4)
            if spacing is None:
                continue

            # Определяем ожидаемый интервал для специальных случаев
            if text.startswith("Условные обозначения:"):
                expected_spacing = 1.0
            else:
                expected_spacing = body_spacing

            if abs(spacing - expected_spacing) > tol:
                if i - last_issue_para >= 5:
                    add_issue(
                        result,
                        rule_id="line_spacing",
                        message=(
                            f"Межстрочный интервал {spacing:.2f} "
                            f"(требуется {expected_spacing})"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )
                    last_issue_para = i