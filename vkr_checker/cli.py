"""
CLI на основе Click.

Использование:
  vkr-check document.docx
  vkr-check document.docx --output report.html
  vkr-check document.docx --format html --output report.html
  vkr-check document.docx --config my_rules.yaml
"""
from __future__ import annotations

from pathlib import Path
import sys

import click
from rich.console import Console

from .orchestrator import run_checks
from .reporters.console_reporter import print_report
from .reporters.html_reporter import save_html_report

console = Console()


@click.command()
@click.argument("docx_file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "--output", "-o",
    type=click.Path(path_type=Path),
    default=None,
    help="Путь для сохранения отчёта (HTML)",
)
@click.option(
    "--format", "-f",
    "fmt",
    type=click.Choice(["console", "html", "both"], case_sensitive=False),
    default="console",
    show_default=True,
    help="Формат вывода",
)
@click.option(
    "--config", "-c",
    type=click.Path(exists=True, path_type=Path),
    default=None,
    help="Путь к файлу конфигурации rules.yaml",
)
@click.option(
    "--strict",
    is_flag=True,
    default=False,
    help="Возвращать код ошибки 1 при наличии предупреждений",
)
def main(
    docx_file: Path,
    output: Path | None,
    fmt: str,
    config: Path | None,
    strict: bool,
) -> None:
    """Проверка оформления ВКР по требованиям регламента."""
    console.print(f"[dim]Проверяю: {docx_file}[/dim]")

    try:
        report = run_checks(docx_file, config_path=config)
    except (FileNotFoundError, ValueError) as e:
        console.print(f"[red]Ошибка:[/red] {e}")
        sys.exit(2)

    if fmt in ("console", "both"):
        print_report(report)

    if fmt in ("html", "both"):
        if output is None:
            output = docx_file.with_suffix(".report.html")
        saved = save_html_report(report, output)
        console.print(f"\n[dim]HTML-отчёт сохранён: {saved}[/dim]")

    # Код выхода
    if report.total_errors > 0:
        sys.exit(1)
    if strict and report.total_warnings > 0:
        sys.exit(1)
    sys.exit(0)
