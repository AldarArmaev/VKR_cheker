"""Массовая проверка DOCX-файлов в выбранной папке."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from html import escape
from pathlib import Path
import traceback
from typing import Callable

from .orchestrator import run_checks
from .reporters.html_reporter import save_html_report

ProgressCallback = Callable[[int, int, Path, str], None]


@dataclass
class FileCheckSummary:
    """Краткий результат проверки одного файла."""
    file_path: Path
    status: str
    errors: int | None = None
    warnings: int | None = None
    report_path: Path | None = None
    error_text: str = ""

    @property
    def ok(self) -> bool:
        return self.status == "ok" and (self.errors or 0) == 0


def find_docx_files(folder_path: str | Path) -> list[Path]:
    """Ищет все DOCX-файлы в папке и подпапках, исключая временные файлы Word."""
    folder = Path(folder_path)
    return sorted(
        p for p in folder.rglob("*.docx")
        if p.is_file() and not p.name.startswith("~$")
    )


def check_folder(
    folder_path: str | Path,
    reports_dir: str | Path | None = None,
    progress_callback: ProgressCallback | None = None,
) -> tuple[list[FileCheckSummary], Path, Path]:
    """
    Проверяет все DOCX-файлы в папке.

    Returns:
        summaries: список кратких результатов
        summary_report_path: путь к общему HTML-отчёту
        error_log_path: путь к журналу ошибок
    """
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Папка не найдена: {folder}")

    docx_files = find_docx_files(folder)
    if reports_dir is None:
        reports_dir = folder / "vkr_reports"
    reports_dir = Path(reports_dir)
    reports_dir.mkdir(parents=True, exist_ok=True)

    error_log_path = reports_dir / "errors.log"
    summaries: list[FileCheckSummary] = []

    # Начинаем новый лог для каждого запуска.
    error_log_path.write_text(
        f"VKR Checker — журнал ошибок\nЗапуск: {datetime.now():%d.%m.%Y %H:%M:%S}\n\n",
        encoding="utf-8",
    )

    total = len(docx_files)
    for index, docx_file in enumerate(docx_files, start=1):
        if progress_callback:
            progress_callback(index, total, docx_file, "Проверяется")

        try:
            report = run_checks(docx_file)
            safe_stem = _safe_report_name(docx_file, folder)
            report_path = reports_dir / f"{safe_stem}.report.html"
            save_html_report(report, report_path)

            summaries.append(FileCheckSummary(
                file_path=docx_file,
                status="ok",
                errors=report.total_errors,
                warnings=report.total_warnings,
                report_path=report_path,
            ))

            if progress_callback:
                progress_callback(index, total, docx_file, "Готово")

        except Exception as exc:  # noqa: BLE001 — нужно не падать на одном плохом файле
            error_text = f"{type(exc).__name__}: {exc}"
            summaries.append(FileCheckSummary(
                file_path=docx_file,
                status="failed",
                error_text=error_text,
            ))
            with error_log_path.open("a", encoding="utf-8") as log:
                log.write(f"Файл: {docx_file}\n")
                log.write(error_text + "\n")
                log.write(traceback.format_exc())
                log.write("\n" + "-" * 80 + "\n")

            if progress_callback:
                progress_callback(index, total, docx_file, "Ошибка")

    summary_report_path = reports_dir / "summary_report.html"
    save_summary_report(folder, summaries, summary_report_path, error_log_path)
    return summaries, summary_report_path, error_log_path


def save_summary_report(
    folder: Path,
    summaries: list[FileCheckSummary],
    output_path: Path,
    error_log_path: Path,
) -> Path:
    """Создаёт общий HTML-отчёт по всем проверенным файлам."""
    total_files = len(summaries)
    failed_files = sum(1 for s in summaries if s.status == "failed")
    files_with_errors = sum(1 for s in summaries if s.status == "ok" and (s.errors or 0) > 0)
    ok_files = sum(1 for s in summaries if s.status == "ok" and (s.errors or 0) == 0)
    total_errors = sum(s.errors or 0 for s in summaries)
    total_warnings = sum(s.warnings or 0 for s in summaries)

    rows = []
    for item in summaries:
        if item.status == "failed":
            status = '<span class="bad">Ошибка проверки</span>'
            errors = "—"
            warnings = "—"
            report_link = escape(item.error_text)
        else:
            status = '<span class="good">OK</span>' if (item.errors or 0) == 0 else '<span class="bad">Есть ошибки</span>'
            errors = str(item.errors or 0)
            warnings = str(item.warnings or 0)
            report_link = f'<a href="{escape(item.report_path.name)}">Открыть отчёт</a>' if item.report_path else "—"

        rows.append(f"""
        <tr>
            <td>{escape(str(item.file_path.relative_to(folder)))}</td>
            <td>{status}</td>
            <td>{errors}</td>
            <td>{warnings}</td>
            <td>{report_link}</td>
        </tr>
        """)

    if not rows:
        rows.append("""
        <tr>
            <td colspan="5">В выбранной папке DOCX-файлы не найдены.</td>
        </tr>
        """)

    html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Сводный отчёт VKR Checker</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; max-width: 1100px; margin: 35px auto; padding: 0 20px; color: #222; }}
  h1 {{ border-bottom: 2px solid #333; padding-bottom: 8px; }}
  .summary {{ display: flex; flex-wrap: wrap; gap: 12px; margin: 18px 0; }}
  .card {{ padding: 12px 18px; border-radius: 8px; background: #f2f2f2; min-width: 145px; }}
  .card strong {{ display: block; font-size: 1.5em; }}
  table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
  th, td {{ border: 1px solid #ddd; padding: 9px 10px; text-align: left; vertical-align: top; }}
  th {{ background: #f7f7f7; }}
  .good {{ color: #1e8449; font-weight: bold; }}
  .bad {{ color: #c0392b; font-weight: bold; }}
  .muted {{ color: #777; }}
</style>
</head>
<body>
<h1>Сводный отчёт VKR Checker</h1>
<p><strong>Папка:</strong> {escape(str(folder))}</p>
<p><strong>Дата проверки:</strong> {datetime.now():%d.%m.%Y %H:%M:%S}</p>

<div class="summary">
  <div class="card"><strong>{total_files}</strong> файлов проверено</div>
  <div class="card"><strong>{ok_files}</strong> без ошибок</div>
  <div class="card"><strong>{files_with_errors}</strong> с ошибками</div>
  <div class="card"><strong>{failed_files}</strong> не проверено</div>
  <div class="card"><strong>{total_errors}</strong> ошибок</div>
  <div class="card"><strong>{total_warnings}</strong> предупреждений</div>
</div>

<p class="muted">Журнал технических ошибок: <a href="{escape(error_log_path.name)}">errors.log</a></p>

<table>
<thead>
<tr>
  <th>Файл</th>
  <th>Статус</th>
  <th>Ошибок</th>
  <th>Предупреждений</th>
  <th>Отчёт</th>
</tr>
</thead>
<tbody>
{''.join(rows)}
</tbody>
</table>
</body>
</html>
"""
    output_path.write_text(html, encoding="utf-8")
    return output_path


def _safe_report_name(docx_file: Path, root_folder: Path) -> str:
    """Делает уникальное безопасное имя отчёта из относительного пути файла."""
    rel = docx_file.relative_to(root_folder).with_suffix("")
    return "__".join(_clean_part(part) for part in rel.parts)


def _clean_part(text: str) -> str:
    forbidden = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in forbidden else ch for ch in text)
    return cleaned.strip() or "document"
