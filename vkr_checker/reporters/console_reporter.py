"""Вывод отчёта в консоль с использованием rich."""
from __future__ import annotations

from rich.console import Console
from rich.table import Table
from rich.text import Text
from rich import box

from ..orchestrator import Report
from ..checks.base import Severity


console = Console()

SEVERITY_COLORS = {
    Severity.ERROR:   "red",
    Severity.WARNING: "yellow",
    Severity.INFO:    "blue",
}

SEVERITY_ICONS = {
    Severity.ERROR:   "✗",
    Severity.WARNING: "⚠",
    Severity.INFO:    "ℹ",
}


def print_report(report: Report) -> None:
    """Выводит полный отчёт в консоль."""
    console.print()
    _print_header(report)
    _print_summary_table(report)
    _print_details(report)
    console.print()


def _print_header(report: Report) -> None:
    console.rule(f"[bold]Проверка ВКР: {report.document_name}[/bold]")
    verdict_color = "green" if report.total_errors == 0 else "red"
    console.print(f"\n  {report.verdict}", style=f"bold {verdict_color}")
    console.print(
        f"  Ошибок: [red]{report.total_errors}[/red]  "
        f"Предупреждений: [yellow]{report.total_warnings}[/yellow]  "
        f"Проверок пройдено: [green]{report.checks_passed}/{report.checks_total}[/green]"
    )
    console.print()


def _print_summary_table(report: Report) -> None:
    tbl = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    tbl.add_column("Проверка", min_width=30)
    tbl.add_column("Статус", width=12)
    tbl.add_column("Ошибок", justify="center", width=8)
    tbl.add_column("Предупр.", justify="center", width=10)

    for r in report.results:
        if r.skipped:
            status = Text("пропущено", style="dim")
        elif r.passed:
            status = Text("✓ OK", style="green")
        else:
            status = Text("✗ ОШИБКИ", style="red")

        tbl.add_row(
            r.check_name,
            status,
            str(r.error_count) if r.error_count else "",
            str(r.warning_count) if r.warning_count else "",
        )

    console.print(tbl)


def _print_details(report: Report) -> None:
    has_issues = any(r.issues for r in report.results)
    if not has_issues:
        console.print("[green]  Замечаний не обнаружено.[/green]")
        return

    for result in report.results:
        if not result.issues:
            continue
        console.print(f"\n[bold]{result.check_name}[/bold]")
        for issue in result.issues:
            color = SEVERITY_COLORS[issue.severity]
            icon = SEVERITY_ICONS[issue.severity]
            line = f"  [{color}]{icon}[/{color}] {issue.message}"
            if issue.location_hint:
                line += f"  [dim]{issue.location_hint}[/dim]"
            console.print(line)
            if issue.context:
                console.print(f"      [dim italic]→ «{issue.context}»[/dim italic]")
