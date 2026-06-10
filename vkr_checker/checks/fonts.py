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

        # Компилируем паттерны заголовков, где разрешён жирный шрифт
        bold_patterns = self.rules.get("headings", {}).get("bold_allowed_patterns", [])
        self._bold_patterns = [re.compile(p) for p in bold_patterns]

        # Компилируем паттерны специальных параграфов
        specials = self.rules.get("special_paragraphs", [])
        self._special_patterns = [(re.compile(s["pattern"]), s) for s in specials]

        # Если Введение не найдено, проверять весь документ
        if model.intro_start_idx < 0:
            start_idx = 0
        else:
            start_idx = model.intro_start_idx

        last_font_issue_para = -10
        last_size_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            if i < start_idx:
                continue
            if para_is_in_table(para):
                continue

            text = para.text.strip()
            if not text:
                continue

            # Определяем, является ли параграф специальным
            special_rule = self._match_special_paragraph(text)
            required_size_current = required_size
            if special_rule and "font_size" in special_rule:
                required_size_current = special_rule["font_size"]

            for run in para.runs:
                if not run.text.strip():
                    continue

                font_name = resolver.get_font_name(run, para)
                font_size = resolver.get_font_size_pt(run, para)

                # Проверка шрифта
                if font_name is None:
                    if i - last_font_issue_para >= 3:
                        add_issue(
                            result,
                            rule_id="font_undefined",
                            message="Шрифт не определён (возможно, наследуется из темы)",
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_font_issue_para = i
                elif font_name != required_font:
                    if i - last_font_issue_para >= 3:
                        add_issue(
                            result,
                            rule_id="font_name",
                            message=(
                                f"Шрифт «{font_name}» "
                                f"(требуется «{required_font}»)"
                            ),
                            severity=Severity.ERROR,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_font_issue_para = i
                    break

                # Проверка кегля
                if font_size and abs(font_size - required_size_current) > 0.5:
                    if i - last_size_issue_para >= 3:
                        add_issue(
                            result,
                            rule_id="font_size",
                            message=(
                                f"Кегль {font_size:.0f} пт "
                                f"(требуется {required_size_current} пт)"
                            ),
                            severity=Severity.ERROR,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_size_issue_para = i
                    break

        # Проверка жирного шрифта в тексте (запрещён вне заголовков)
        self._check_bold_in_body(model, resolver, result)

    def _match_special_paragraph(self, text: str) -> dict | None:
        """Возвращает правила для специального параграфа, если текст соответствует паттерну."""
        for pattern, rule in self._special_patterns:
            if pattern.match(text):
                return rule
        return None

    def _check_bold_in_body(self, model, resolver, result: CheckResult) -> None:
        in_intro = False
        last_bold_issue = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            if para_is_in_table(para):
                continue

            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_intro = True
            elif re.match(r"^Глава\s+1", text, re.IGNORECASE):
                in_intro = False

            if self._is_heading_paragraph(para):
                continue
            if in_intro:
                continue

            for run in para.runs:
                if resolver.is_bold(run, para) and run.text.strip():
                    if i - last_bold_issue >= 5:
                        add_issue(
                            result,
                            rule_id="bold_in_body",
                            message="Жирный шрифт в основном тексте (запрещён)",
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_bold_issue = i
                    break

    def _is_heading_paragraph(self, para) -> bool:
        text = para.text.strip()
        return any(p.match(text) for p in self._bold_patterns)