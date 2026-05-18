"""
Проверка шрифта и кегля.

Алгоритм:
  - Для каждого run в основном тексте разрезолвить шрифт через StyleResolver.
  - Сравнить с требованиями.
  - Пропускать runs внутри таблиц (их проверяет tables.py).
  - Пропускать пустые runs.
  - Дедуплицировать: одна ошибка на каждые 5 подряд нарушающих параграфов.
"""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue


# Заголовки, где разрешён жирный шрифт
BOLD_ALLOWED_PATTERNS = [
    re.compile(r"^Глава\s+\d+"),
    re.compile(r"^\d+\.\d+\."),
    re.compile(r"^Выводы по главе"),
    re.compile(r"^Введение$", re.IGNORECASE),
    re.compile(r"^Заключение$", re.IGNORECASE),
    re.compile(r"^Список литературы$", re.IGNORECASE),
    re.compile(r"^Содержание$", re.IGNORECASE),
    # Подразделы Введения
    re.compile(r"^(Актуальность|Цель|Объект|Предмет|Гипотез|Задачи|"
               r"Теоретическая|Общая характеристика|Методы|Структура)"),
]


def is_heading_paragraph(para) -> bool:
    text = para.text.strip()
    return any(p.match(text) for p in BOLD_ALLOWED_PATTERNS)


def para_is_in_table(para) -> bool:
    """Проверяет, находится ли параграф внутри таблицы."""
    parent = para._p.getparent()
    while parent is not None:
        if parent.tag.endswith("}tc"):  # w:tc — ячейка таблицы
            return True
        parent = parent.getparent()
    return False


class FontsCheck(BaseCheck):
    check_id = "fonts"
    check_name = "Шрифт и кегль"

    def _run(self, model, resolver, result: CheckResult) -> None:
        required_font = self.rules["fonts"]["required_font"]
        required_size = self.rules["fonts"]["sizes"]["body"]

        # Проверяем наличие раздела «Введение» — без него проверки невозможны
        if model.intro_start_idx < 0:
            result.skipped = True
            result.skip_reason = (
                "Раздел «Введение» не найден. "
                "Проверка шрифтов требует наличия этого раздела."
            )
            return

        wrong_font_count = 0
        wrong_size_count = 0
        last_font_issue_para = -10
        last_size_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            if not para.text.strip():
                continue
            if para_is_in_table(para):
                continue  # таблицы проверяются отдельно

            for run in para.runs:
                if not run.text.strip():
                    continue

                font_name = resolver.get_font_name(run, para)
                font_size = resolver.get_font_size_pt(run, para)

                # Проверка шрифта
                # Если font_name is None (код темы), считаем это нарушением
                if font_name is None:
                    # Шрифт не определён или код темы — предупреждение
                    if i - last_font_issue_para >= 3:
                        add_issue(
                            result,
                            rule_id="font_undefined",
                            message="Шрифт не определён (возможно, наследуется из темы)",
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=para.text[:80],
                        )
                        last_font_issue_para = i
                elif font_name != required_font:
                    wrong_font_count += 1
                    if i - last_font_issue_para >= 3:  # дедупликация
                        add_issue(
                            result,
                            rule_id="font_name",
                            message=(
                                f"Шрифт «{font_name}» "
                                f"(требуется «{required_font}»)"
                            ),
                            severity=Severity.ERROR,
                            location_hint=f"~абз. {i+1}",
                            context=para.text[:80],
                        )
                        last_font_issue_para = i
                    break  # один issue на параграф

                # Проверка кегля (только для основного текста)
                if font_size and abs(font_size - required_size) > 0.5:
                    wrong_size_count += 1
                    if i - last_size_issue_para >= 3:
                        add_issue(
                            result,
                            rule_id="font_size",
                            message=(
                                f"Кегль {font_size:.0f} пт "
                                f"(требуется {required_size} пт)"
                            ),
                            severity=Severity.ERROR,
                            location_hint=f"~абз. {i+1}",
                            context=para.text[:80],
                        )
                        last_size_issue_para = i
                    break

        # Проверка жирного шрифта в тексте (запрещён вне заголовков)
        self._check_bold_in_body(model, resolver, result)

    def _check_bold_in_body(self, model, resolver, result: CheckResult) -> None:
        """Жирный шрифт запрещён в основном тексте, кроме заголовков и Введения."""
        in_intro = False
        last_bold_issue = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            if para_is_in_table(para):
                continue

            # Отслеживаем вход/выход из Введения
            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_intro = True
            elif re.match(r"^Глава\s+1", text, re.IGNORECASE):
                in_intro = False

            # В заголовках жирный разрешён
            if is_heading_paragraph(para):
                continue
            # В подразделах Введения жирный разрешён
            if in_intro:
                continue

            for run in para.runs:
                if run.bold and run.text.strip():
                    if i - last_bold_issue >= 5:
                        add_issue(
                            result,
                            rule_id="bold_in_body",
                            message="Жирный шрифт в основном тексте (запрещён)",
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=para.text[:80],
                        )
                        last_bold_issue = i
                    break
