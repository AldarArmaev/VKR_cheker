"""Оркестратор: загружает документ, запускает все проверки, возвращает итоговый отчёт."""
from __future__ import annotations

import sys
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
from .checks.heading_styles import HeadingStylesCheck
from .checks.sections import SectionsCheck
from .checks.intro import IntroCheck
from .checks.tables import TablesCheck

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
    model: DocumentModel = load_document(docx_path, rules)
    resolver = StyleResolver(model.doc)

    # Порядок важен: сначала структурные, потом детальные
    check_classes = [
        MarginsCheck,
        HeadingStylesCheck,
        SectionsCheck,  # проверка наличия разделов
        IntroCheck,  # проверка структуры Введения
        FontsCheck,
        SpacingCheck,
        IndentsCheck,
        TablesCheck,  # опционально
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


def print_report(report: Report) -> None:
    """Выводит отчёт в консоль."""
    print("\n" + "=" * 60)
    print(f"Отчёт по документу: {report.document_name}")
    print("=" * 60)
    
    for result in report.results:
        if result.skipped:
            print(f"\n[{result.check_name}] ПРОПУЩЕНО: {result.skip_reason}")
            continue
        
        status = "✓" if result.passed else "✗"
        print(f"\n[{result.check_name}] {status}")
        
        if result.issues:
            for issue in result.issues:
                print(f"  {issue}")
    
    print("\n" + "=" * 60)
    print(f"Проверок пройдено: {report.checks_passed}/{report.checks_total}")
    print(f"Всего ошибок: {report.total_errors}")
    print(f"Всего предупреждений: {report.total_warnings}")
    print(f"Вердикт: {report.verdict}")
    print("=" * 60 + "\n")


def main() -> int:
    """Точка входа CLI."""
    if len(sys.argv) < 2:
        print("Использование: python -m vkr_checker <путь_к_файлу.docx>")
        print("Пример: python -m vkr_checker diploma.docx")
        return 1
    
    docx_path = sys.argv[1]
    
    try:
        report = run_checks(docx_path)
        print_report(report)
        return 0 if report.total_errors == 0 else 1
    except FileNotFoundError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 2
    except ValueError as e:
        print(f"Ошибка формата: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"Неожиданная ошибка: {e}", file=sys.stderr)
        return 4


if __name__ == "__main__":
    sys.exit(main())
