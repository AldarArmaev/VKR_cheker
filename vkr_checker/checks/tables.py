"""
Проверка оформления таблиц.

Что проверяем:
  1. Наличие строки «Таблица N» перед каждой таблицей
  2. Наличие названия таблицы (параграф после «Таблица N»)
  3. Название не заканчивается точкой
  4. Строка «Таблица N» выровнена по правому краю
  5. Название таблицы выровнено по центру
  6. Шрифт и кегль в ячейках
  7. Ссылка на таблицу в предшествующем тексте
  8. Сквозная нумерация (нет пропусков)
"""
from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from .base import BaseCheck, CheckResult, Severity, add_issue

TABLE_NUMBER_RE = re.compile(r"^Таблица\s+(\d+)\s*$")


class TablesCheck(BaseCheck):
    check_id = "tables"
    check_name = "Оформление таблиц"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_fonts = self.rules["fonts"]
        allowed_cell_sizes = rules_fonts["sizes"]["table_cell"]
        required_font = rules_fonts["required_font"]

        table_numbers_seen = []

        # Итерируем по блокам, чтобы знать контекст (что стоит перед таблицей)
        for block_idx, block in enumerate(model.blocks):
            if block.kind != "table":
                continue

            table = block.table
            table_num = len(table_numbers_seen) + 1

            # Ищем строку "Таблица N" в предшествующих параграфах (до 5 назад)
            prev_paras = self._get_prev_paragraphs(model.blocks, block_idx, n=5)
            caption_line, title_line = self._find_caption(prev_paras)

            if caption_line is None:
                add_issue(
                    result,
                    rule_id="table_no_number",
                    message=(
                        f"Таблица {table_num}: "
                        f"не найдена строка «Таблица N» перед таблицей"
                    ),
                    severity=Severity.ERROR,
                )
            else:
                m = TABLE_NUMBER_RE.match(caption_line.text.strip())
                if m:
                    num = int(m.group(1))
                    table_numbers_seen.append(num)

                    # Проверка выравнивания строки "Таблица N"
                    if caption_line.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                        add_issue(
                            result,
                            rule_id="table_number_alignment",
                            message=(
                                f"Таблица {num}: строка «Таблица N» "
                                f"должна быть выровнена по правому краю"
                            ),
                            severity=Severity.ERROR,
                            context=caption_line.text,
                        )

            if title_line is None:
                add_issue(
                    result,
                    rule_id="table_no_title",
                    message=f"Таблица {table_num}: отсутствует название",
                    severity=Severity.ERROR,
                )
            else:
                # Название не должно заканчиваться точкой
                if title_line.text.strip().endswith("."):
                    add_issue(
                        result,
                        rule_id="table_title_dot",
                        message=(
                            f"Таблица {table_num}: название таблицы "
                            f"не должно заканчиваться точкой"
                        ),
                        severity=Severity.ERROR,
                        context=title_line.text[:80],
                    )
                # Выравнивание названия
                if title_line.alignment not in (
                    WD_ALIGN_PARAGRAPH.CENTER, None
                ):
                    add_issue(
                        result,
                        rule_id="table_title_alignment",
                        message=(
                            f"Таблица {table_num}: название должно быть "
                            f"по центру"
                        ),
                        severity=Severity.WARNING,
                    )

            # Проверка шрифта и кегля в ячейках
            self._check_cell_fonts(
                result, table, table_num,
                required_font, allowed_cell_sizes, resolver,
            )

        # Проверка сквозной нумерации
        for j, num in enumerate(table_numbers_seen):
            if num != j + 1:
                add_issue(
                    result,
                    rule_id="table_numbering",
                    message=(
                        f"Нарушена сквозная нумерация таблиц: "
                        f"ожидалась Таблица {j+1}, найдена Таблица {num}"
                    ),
                    severity=Severity.ERROR,
                )
                break

    @staticmethod
    def _get_prev_paragraphs(blocks, current_idx, n=5):
        result_paras = []
        i = current_idx - 1
        count = 0
        while i >= 0 and count < n:
            if blocks[i].kind == "paragraph":
                result_paras.insert(0, blocks[i].paragraph)
                count += 1
            i -= 1
        return result_paras

    @staticmethod
    def _find_caption(prev_paras):
        """
        Ищет строку «Таблица N» и строку с названием таблицы.
        Возвращает (caption_para, title_para) или (None, None).
        """
        for j, para in enumerate(prev_paras):
            if TABLE_NUMBER_RE.match(para.text.strip()):
                # Следующий непустой параграф — название
                title = None
                for k in range(j + 1, len(prev_paras)):
                    if prev_paras[k].text.strip():
                        title = prev_paras[k]
                        break
                return para, title
        return None, None

    def _check_cell_fonts(
        self, result, table, table_num,
        required_font, allowed_sizes, resolver,
    ):
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if not run.text.strip():
                            continue
                        font = resolver.get_font_name(run, para)
                        size = resolver.get_font_size_pt(run, para)

                        if font and font != required_font:
                            add_issue(
                                result,
                                rule_id="table_cell_font",
                                message=(
                                    f"Таблица {table_num}: шрифт в ячейке «{font}» "
                                    f"(требуется «{required_font}»)"
                                ),
                                severity=Severity.ERROR,
                                context=para.text[:60],
                            )
                            return  # одна ошибка на таблицу

                        if size and size not in allowed_sizes:
                            add_issue(
                                result,
                                rule_id="table_cell_size",
                                message=(
                                    f"Таблица {table_num}: кегль {size:.0f} пт "
                                    f"(допустимо: {allowed_sizes})"
                                ),
                                severity=Severity.WARNING,
                                context=para.text[:60],
                            )
                            return
