"""Проверка выравнивания текста."""
from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import para_is_in_table
from docx.oxml.ns import qn

ALIGN_NAMES = {
    WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширине",
    WD_ALIGN_PARAGRAPH.CENTER:  "по центру",
    WD_ALIGN_PARAGRAPH.RIGHT:   "по правому краю",
    WD_ALIGN_PARAGRAPH.LEFT:    "по левому краю",
    None:                        "по левому краю (по умолчанию)",
}

ALIGN_CODES = {
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
}

TABLE_NUMBER_RE = re.compile(r"^Таблица\s+\d+\s*$")
FIGURE_CAPTION_RE = re.compile(r"^Рис\.\s+\d+\.")


class AlignmentCheck(BaseCheck):
    check_id = "alignment"
    check_name = "Выравнивание"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["alignment"]
        body_align = ALIGN_CODES[rules["body"]]
        chapter_align = ALIGN_CODES[rules["chapter_heading"]]
        table_num_align = ALIGN_CODES[rules["table_number_line"]]
        table_cap_align = ALIGN_CODES[rules["table_caption_title"]]
        fig_cap_align = ALIGN_CODES[rules["figure_caption"]]

        in_main_text = False
        last_issue = -10
        # После строки "Таблица N" следующий параграф — название таблицы
        next_is_table_title = False

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True
            if not in_main_text or not text:
                next_is_table_title = False
                continue
            if para_is_in_table(para):
                next_is_table_title = False
                continue

            actual_align_str = resolver.get_alignment(para)
            actual_align = ALIGN_CODES.get(actual_align_str)

            # Строка "Таблица N"
            if TABLE_NUMBER_RE.match(text):
                self._check_align(result, i, para, actual_align,
                                  table_num_align, "строка «Таблица N»")
                next_is_table_title = True
                continue

            # Название таблицы
            if next_is_table_title:
                self._check_align(result, i, para, actual_align,
                                  table_cap_align, "название таблицы")
                next_is_table_title = False
                continue

            next_is_table_title = False

            # Подпись рисунка
            if FIGURE_CAPTION_RE.match(text):
                self._check_align(result, i, para, actual_align,
                                  fig_cap_align, "подпись рисунка")
                continue

            # Заголовки глав/параграфов
            if is_heading_for_alignment(text, para, resolver):
                self._check_align(result, i, para, actual_align,
                                  chapter_align, "заголовок")
                continue

            # Основной текст (дедупликация)
            if actual_align not in (body_align, None):  # None = left по умолчанию
                if actual_align != WD_ALIGN_PARAGRAPH.LEFT:  # left = нарушение
                    if i - last_issue >= 5:
                        add_issue(
                            result,
                            rule_id="body_alignment",
                            message=(
                                f"Выравнивание: {ALIGN_NAMES.get(actual_align)} "
                                f"(требуется по ширине)"
                            ),
                            severity=Severity.WARNING,
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_issue = i

    @staticmethod
    def _check_align(result, i, para, actual, expected, label):
        if actual != expected and not (
            actual is None and expected == WD_ALIGN_PARAGRAPH.LEFT
        ):
            add_issue(
                result,
                rule_id=f"alignment_{label.replace(' ', '_')}",
                message=(
                    f"{label.capitalize()}: выравнивание «{ALIGN_NAMES.get(actual)}» "
                    f"(требуется «{ALIGN_NAMES.get(expected)}»)"
                ),
                severity=Severity.ERROR,
                location_hint=f"~абз. {i+1}",
                context=para.text[:80],
            )

def is_real_heading_for_alignment(text: str, para) -> bool:
    if re.match(r"^Глава\s+\d+", text):
        return True
    if re.match(r"^\d+\.\d+(\.\d+)*\.", text):
        return True
    if re.match(r"^Выводы по главе", text):
        return True
    if re.match(r"^Введение$", text, re.IGNORECASE):
        return True
    if re.match(r"^Заключение$", text, re.IGNORECASE):
        return True
    if re.match(r"^Список литературы$", text, re.IGNORECASE):
        return True
    if re.match(r"^Содержание$", text, re.IGNORECASE):
        return True
    return False


def is_numbered_paragraph(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


def is_heading_for_alignment(text: str, para, resolver) -> bool:
    if is_real_heading_for_alignment(text):
        return True

    if is_numbered_paragraph(para) and resolver.get_alignment(para) == "center":
        return True

    return False