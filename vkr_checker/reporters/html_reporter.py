"""Генерация HTML-отчёта."""
from __future__ import annotations

from pathlib import Path
from jinja2 import Template
from ..orchestrator import Report
from ..checks.base import Severity

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Отчёт проверки ВКР — {{ report.document_name }}</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 900px;
         margin: 40px auto; padding: 0 20px; color: #222; }
  h1   { font-size: 1.4em; border-bottom: 2px solid #333; padding-bottom: 8px; }
  .verdict-ok  { color: #2a7a2a; font-weight: bold; }
  .verdict-err { color: #c0392b; font-weight: bold; }
  .summary { display: flex; gap: 24px; margin: 16px 0; }
  .stat { padding: 10px 18px; border-radius: 6px; font-size: 1.1em; }
  .stat-errors   { background: #fde8e8; color: #c0392b; }
  .stat-warnings { background: #fef9e7; color: #b7770d; }
  .stat-passed   { background: #eafaf1; color: #1e8449; }
  .check-section { margin-top: 20px; }
  .check-title   { font-size: 1em; font-weight: bold; margin-bottom: 6px;
                   cursor: pointer; }
  .check-ok   .check-title { color: #1e8449; }
  .check-err  .check-title { color: #c0392b; }
  .check-warn .check-title { color: #b7770d; }
  .issue      { padding: 5px 10px; margin: 3px 0; border-radius: 4px;
                font-size: 0.93em; }
  .issue-error   { background: #fde8e8; border-left: 3px solid #c0392b; }
  .issue-warning { background: #fef9e7; border-left: 3px solid #e67e22; }
  .issue-info    { background: #eaf4fb; border-left: 3px solid #2980b9; }
  .location { color: #888; font-size: 0.9em; margin-left: 8px; }
  .context  { color: #555; font-style: italic; font-size: 0.88em;
              margin-top: 2px; }
</style>
</head>
<body>
<h1>Отчёт проверки ВКР</h1>
<p><strong>Файл:</strong> {{ report.document_name }}</p>

<p class="{{ 'verdict-ok' if report.total_errors == 0 else 'verdict-err' }}">
  {{ report.verdict }}
</p>

<div class="summary">
  <div class="stat stat-errors">Ошибок: {{ report.total_errors }}</div>
  <div class="stat stat-warnings">Предупреждений: {{ report.total_warnings }}</div>
  <div class="stat stat-passed">Проверок OK: {{ report.checks_passed }}/{{ report.checks_total }}</div>
</div>

{% for result in report.results %}
  {% if not result.skipped %}
  <div class="check-section {{ 'check-ok' if result.passed else ('check-warn' if result.error_count == 0 else 'check-err') }}">
    <div class="check-title">
      {{ '✓' if result.passed else '✗' }} {{ result.check_name }}
      {% if result.issues %}({{ result.issues|length }}){% endif %}
    </div>
    {% for issue in result.issues %}
    <div class="issue issue-{{ issue.severity.value }}">
      {{ issue.message }}
      {% if issue.location_hint %}
        <span class="location">{{ issue.location_hint }}</span>
      {% endif %}
      {% if issue.context %}
        <div class="context">→ «{{ issue.context }}»</div>
      {% endif %}
    </div>
    {% endfor %}
  </div>
  {% endif %}
{% endfor %}

<p style="margin-top:40px; color:#aaa; font-size:0.85em;">
  Сформировано: vkr-checker
</p>
</body>
</html>
"""


def save_html_report(report: Report, output_path: str | Path) -> Path:
    """Сохраняет HTML-отчёт в файл."""
    output_path = Path(output_path)
    tmpl = Template(HTML_TEMPLATE)
    html = tmpl.render(report=report, Severity=Severity)
    output_path.write_text(html, encoding="utf-8")
    return output_path
