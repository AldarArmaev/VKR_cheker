"""Модуль проверок формальных параметров ВКР."""
from .base import BaseCheck, CheckResult, Issue, Severity, add_issue
from .sections import SectionsCheck
from .intro import IntroCheck
from .page_numbers import PageNumbersCheck

__all__ = [
    "BaseCheck",
    "CheckResult",
    "Issue",
    "Severity",
    "add_issue",
    "SectionsCheck",
    "IntroCheck",
    "PageNumbersCheck",
]
