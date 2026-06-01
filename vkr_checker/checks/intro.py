"""Проверка структуры раздела Введение."""
from __future__ import annotations

from .base import BaseCheck, CheckResult, Severity, add_issue


class IntroCheck(BaseCheck):
    """Проверка структуры и оформления подразделов Введения."""

    check_id = "intro"
    check_name = "Структура Введения"

    def _run(self, model, resolver, result: CheckResult) -> None:
        if model.intro_start_idx < 0:
            add_issue(
                result,
                rule_id="intro_not_found",
                message="Раздел «Введение» не найден — проверка структуры невозможна",
                severity=Severity.ERROR,
            )
            return

        subsections = self.rules["intro_subsections"]
        intro_paras = model.paragraphs[
            model.intro_start_idx : model.intro_end_idx
        ]
        intro_text = " ".join(p.text.strip() for p in intro_paras)

        for sub in subsections:
            keyword = sub["keyword"]
            bold_required = sub.get("bold_required", True)

            # Ищем параграф, начинающийся с ключевого слова
            found_para = None
            for para in intro_paras:
                if para.text.strip().startswith(keyword):
                    found_para = para
                    break

            if found_para is None:
                add_issue(
                    result,
                    rule_id=f"intro_missing_{keyword.replace(' ', '_')[:20]}",
                    message=f"Во Введении отсутствует подраздел «{keyword}»",
                    severity=Severity.ERROR,
                )
                continue

            # Проверяем, что ключевое слово выделено жирным
            if bold_required:
                is_bold = any(
                    run.bold and keyword[:6] in run.text
                    for run in found_para.runs
                )
                if not is_bold:
                    add_issue(
                        result,
                        rule_id=f"intro_bold_{keyword.replace(' ', '_')[:20]}",
                        message=(
                            f"«{keyword}» во Введении должно быть "
                            f"выделено жирным шрифтом"
                        ),
                        severity=Severity.WARNING,
                        context=found_para.text[:80],
                    )
