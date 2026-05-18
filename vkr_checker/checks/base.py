"""
Базовые классы для системы проверок.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Severity(Enum):
    ERROR = "error"       # Грубое нарушение — нормоконтроль не пропустит
    WARNING = "warning"   # Вероятная ошибка — стоит проверить вручную
    INFO = "info"         # Информационное замечание


SEVERITY_LABELS = {
    Severity.ERROR:   "✗ ОШИБКА",
    Severity.WARNING: "⚠ ПРЕДУПРЕЖДЕНИЕ",
    Severity.INFO:    "ℹ ИНФОРМАЦИЯ",
}


@dataclass
class Issue:
    """Одно обнаруженное замечание."""
    rule_id: str             # уникальный идентификатор правила
    message: str             # человекочитаемое описание
    severity: Severity = Severity.ERROR
    location_hint: str = ""  # например: "~стр. 12" или "Таблица 3"
    context: str = ""        # цитата проблемного текста (до 80 символов)

    def __str__(self) -> str:
        parts = [SEVERITY_LABELS[self.severity], self.message]
        if self.location_hint:
            parts.append(f"({self.location_hint})")
        if self.context:
            parts.append(f"→ «{self.context[:80]}»")
        return "  ".join(parts)


@dataclass
class CheckResult:
    """Результат одного модуля проверок."""
    check_id: str
    check_name: str
    issues: list[Issue] = field(default_factory=list)
    skipped: bool = False
    skip_reason: str = ""

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.WARNING)

    @property
    def passed(self) -> bool:
        return self.error_count == 0 and not self.skipped


class BaseCheck:
    """
    Базовый класс для всех модулей проверок.

    Каждый наследник обязан:
    - задать class-атрибуты check_id и check_name
    - реализовать метод _run(model, resolver, rules)
    """
    check_id: str = "base"
    check_name: str = "Базовая проверка"

    def __init__(self, rules: dict):
        self.rules = rules

    def run(self, model, resolver) -> CheckResult:
        """Запускает проверку и возвращает CheckResult."""
        result = CheckResult(
            check_id=self.check_id,
            check_name=self.check_name,
        )
        try:
            self._run(model, resolver, result)
        except Exception as exc:
            result.skipped = True
            result.skip_reason = f"Ошибка при выполнении проверки: {exc}"
        return result

    def _run(self, model, resolver, result: CheckResult) -> None:
        raise NotImplementedError


def add_issue(
    result: CheckResult,
    rule_id: str,
    message: str,
    severity: Severity = Severity.ERROR,
    location_hint: str = "",
    context: str = "",
) -> None:
    """Вспомогательная функция добавления замечания."""
    result.issues.append(Issue(
        rule_id=rule_id,
        message=message,
        severity=severity,
        location_hint=location_hint,
        context=context[:100],
    ))
