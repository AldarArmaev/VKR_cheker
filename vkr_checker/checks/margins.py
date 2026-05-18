"""Проверка полей страницы."""
from .base import BaseCheck, CheckResult, Severity, add_issue


class MarginsCheck(BaseCheck):
    check_id = "margins"
    check_name = "Поля страницы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        expected = self.rules["margins_cm"]
        tol = expected.get("tolerance", 0.15)

        for i, section in enumerate(model.doc.sections):
            actual = {
                "left":   section.left_margin.cm,
                "right":  section.right_margin.cm,
                "top":    section.top_margin.cm,
                "bottom": section.bottom_margin.cm,
            }
            labels = {
                "left": "левое", "right": "правое",
                "top": "верхнее", "bottom": "нижнее",
            }
            for side, label in labels.items():
                exp_val = expected[side]
                act_val = actual[side]
                if abs(act_val - exp_val) > tol:
                    add_issue(
                        result,
                        rule_id=f"margin_{side}",
                        message=(
                            f"Поле «{label}»: {act_val:.2f} см "
                            f"(требуется {exp_val:.1f} см)"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"Секция {i+1}",
                    )
