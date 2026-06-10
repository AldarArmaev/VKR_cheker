"""Проверка абзацных отступов."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import para_is_in_table
from docx.oxml.ns import qn

class IndentsCheck(BaseCheck):
    check_id = "indents"
    check_name = "Абзацные отступы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["indents"]
        body_indent = rules["body_first_line"]
        tol = rules.get("tolerance", 0.1)

        # Проверяем наличие раздела «Введение» — без него проверки невозможны
        if model.intro_start_idx < 0:
            result.skipped = True
            result.skip_reason = (
                "Раздел «Введение» не найден. "
                "Проверка отступов требует наличия этого раздела."
            )
            return

        in_main_text = False
        in_bibliography = False
        last_issue_para = -10
        next_is_table_title = False   # флаг для названия таблицы
        next_is_appendix_title = False # флаг для названия приложения

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True
            if not in_main_text or not text:
                continue
            if para_is_in_table(para):
                continue

            # Строка "Таблица N" — следующий непустой абзац будет названием таблицы
            if re.match(r"^Таблица\s+\d+\s*$", text):
                next_is_table_title = True
                continue

            # Строка "Приложение N" — следующий непустой абзац будет названием приложения
            if re.match(r"^Приложение\s+\d+\s*$", text, re.IGNORECASE):
                next_is_appendix_title = True
                continue

            if re.match(r"^Список литературы$", text, re.IGNORECASE):
                in_bibliography = True

            if re.match(r"^Приложени", text, re.IGNORECASE):
                in_bibliography = False

            indent = resolver.get_first_line_indent_cm(para)

            # Проверка названия таблицы (без отступа)
            if next_is_table_title:
                if indent > tol:
                    add_issue(
                        result,
                        rule_id="table_title_indent",
                        message=(
                            f"Название таблицы имеет красную строку {indent:.2f} см "
                            f"(должно быть 0)"
                        ),
                        severity=Severity.ERROR,
                        context=text[:80],
                    )
                next_is_table_title = False
                continue

            # Проверка названия приложения (без отступа)
            if next_is_appendix_title:
                if indent > tol:
                    add_issue(
                        result,
                        rule_id="appendix_title_indent",
                        message=(
                            f"Название приложения имеет красную строку {indent:.2f} см "
                            f"(должно быть 0)"
                        ),
                        severity=Severity.ERROR,
                        context=text[:80],
                    )
                next_is_appendix_title = False
                continue

            is_heading = (
                    is_real_heading_for_indent(text)
                    or (
                            not in_bibliography
                            and looks_like_subsection_title(text, para, resolver)
                    )
            )

            if is_heading:
                # Заголовки должны быть БЕЗ отступа
                if indent > tol:
                    add_issue(
                        result,
                        rule_id="heading_indent",
                        message=(
                            f"Заголовок имеет красную строку {indent:.2f} см "
                            f"(должно быть 0)"
                        ),
                        severity=Severity.ERROR,
                        context=text[:80],
                    )
                # Заголовки больше не проверяем, переходим к следующему абзацу
                continue

            # Специальные случаи, которые не должны проверяться как обычный текст
            if text.startswith("Рис."):
                if indent > tol:
                    add_issue(
                        result,
                        rule_id="figure_caption_indent",
                        message=(
                            f"Подпись рисунка имеет красную строку {indent:.2f} см "
                            f"(должно быть 0)"
                        ),
                        severity=Severity.ERROR,
                        context=text[:80],
                    )
                continue

            if text.startswith("Условные обозначения:"):
                # Пояснение к рисунку может быть без красной строки
                continue

            # Основной текст — отступ 1.25 см
            # Исключения из rules.yaml
            matched_exception = False

            for exc in self.rules.get("indent_exceptions", []):
                if re.match(exc["pattern"], text, re.IGNORECASE):

                    expected_indent = exc["expected_indent"]

                    if abs(indent - expected_indent) > tol:
                        add_issue(
                            result,
                            rule_id="indent_exception",
                            message=(
                                f"{exc['label']}: красная строка "
                                f"{indent:.2f} см (требуется {expected_indent:.2f} см)"
                            ),
                            severity=Severity.ERROR,
                            context=text[:80],
                        )

                    matched_exception = True
                    break

            if matched_exception:
                continue
            if abs(indent - body_indent) > tol:
                if i - last_issue_para >= 5:
                    add_issue(
                        result,
                        rule_id="body_indent",
                        message=(
                            f"Красная строка {indent:.2f} см "
                            f"(требуется {body_indent} см)"
                        ),
                        severity=Severity.WARNING,
                        context=text[:80],
                    )
                    last_issue_para = i

REAL_HEADING_PATTERNS = [
    re.compile(r"^Глава\s+\d+"),
    re.compile(r"^\d+\.\d+\."),
    re.compile(r"^Выводы по главе"),
    re.compile(r"^Введение$", re.IGNORECASE),
    re.compile(r"^Заключение$", re.IGNORECASE),
    re.compile(r"^Список литературы$", re.IGNORECASE),
    re.compile(r"^Содержание$", re.IGNORECASE),
]


def is_real_heading_for_indent(text: str) -> bool:
    return any(p.match(text.strip()) for p in REAL_HEADING_PATTERNS)

def is_numbered_paragraph(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    numPr = pPr.find(qn("w:numPr"))
    return numPr is not None

def get_list_level(para):
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    ilvl = numPr.find(qn("w:ilvl"))
    if ilvl is not None and ilvl.get(qn("w:val")):
        return int(ilvl.get(qn("w:val")))
    return None

def looks_like_subsection_title(text: str, para, resolver) -> bool:
    if not is_numbered_paragraph(para):
        return False

    # Уровень 0 (главы) не обрабатываем здесь — они распознаются is_real_heading_for_indent
    # Поэтому удаляем условие if level == 0: return True

    if re.match(r"^\d+\.\d+(\.\d+)*\.\s+", text):
        return True

    return resolver.get_alignment(para) == "center"