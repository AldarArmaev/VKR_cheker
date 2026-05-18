"""Оркестратор: загружает документ, запускает все проверки, возвращает итоговый отчёт."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .parser import load_document, DocumentModel
from .style_resolver import StyleResolver
from .checks.base import CheckResult, Severity
from .checks.margins import MarginsCheck
from .checks.fonts import FontsCheck
from .checks.spacing import SpacingCheck
from .checks.indents import IndentsCheck


@dataclass
class Report:
    """Итоговый отчёт по всему документу."""
    document_path: Path
    document_name: str
    results: list[CheckResult] = field(default_factory=list)

    @property
    def total_errors(self) -> int:
        return sum(r.error_count for r in self.results)

    @property
    def total_warnings(self) -> int:
        return sum(r.warning_count for r in self.results)

    @property
    def checks_passed(self) -> int:
        return sum(1 for r in self.results if r.passed)

    @property
    def checks_total(self) -> int:
        return len([r for r in self.results if not r.skipped])

    @property
    def verdict(self) -> str:
        if self.total_errors == 0:
            return "✓ Формальные требования выполнены"
        return f"✗ Обнаружено {self.total_errors} ошибок"


def load_rules(config_path: str | Path | None = None) -> dict:
    """Загружает конфигурацию требований из YAML."""
    if config_path is None:
        config_path = Path(__file__).parent.parent / "config" / "rules.yaml"
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_checks(
    docx_path: str | Path,
    config_path: str | Path | None = None,
) -> Report:
    """
    Главная функция: загружает документ, запускает все проверки.

    Args:
        docx_path: путь к проверяемому .docx файлу
        config_path: путь к rules.yaml (опционально)

    Returns:
        Report с результатами всех проверок
    """
    rules = load_rules(config_path)
    model: DocumentModel = load_document(docx_path)
    resolver = StyleResolver(model.doc)

    # Порядок важен: сначала структурные, потом детальные
    check_classes = [
        MarginsCheck,
        FontsCheck,
        SpacingCheck,
        IndentsCheck,
    ]

    report = Report(
        document_path=Path(docx_path),
        document_name=Path(docx_path).name,
    )

    for cls in check_classes:
        check = cls(rules)
        result = check.run(model, resolver)
        report.results.append(result)

    return report
