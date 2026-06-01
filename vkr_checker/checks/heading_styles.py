"""Проверка запрещённых стилей заголовков."""
from .base import BaseCheck, CheckResult, Severity, add_issue


class HeadingStylesCheck(BaseCheck):
    check_id = "heading_styles"
    check_name = "Стили заголовков"

    def _run(self, model, resolver, result: CheckResult) -> None:
        forbidden = set(self.rules["styles"]["forbidden"])
        for i, para in enumerate(model.paragraphs):
            if para.style.name in forbidden:
                add_issue(
                    result,
                    "forbidden_style",
                    f"Применён стиль «{para.style.name}» (нужен «Обычный»)",
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                    context=para.text[:80],
                )
