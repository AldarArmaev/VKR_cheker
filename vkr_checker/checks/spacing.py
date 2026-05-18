"""Проверка межстрочного интервала."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import is_heading_paragraph, para_is_in_table


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
            
            # Если интервал не задан (None), берётся из стиля.
            # В типичном ВКР-документе стиль Normal задаёт 1.5,
            # поэтому отсутствие явного значения НЕ является нарушением.
            # Пропускаем такие параграфы без предупреждения.
            if spacing is None:
                continue

            if abs(spacing - body_spacing) > tol:
                if i - last_issue_para >= 5:
                    add_issue(
                        result,
                        rule_id="line_spacing",
                        message=(
                            f"Межстрочный интервал {spacing:.2f} "
                            f"(требуется {body_spacing})"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )
                    last_issue_para = i
