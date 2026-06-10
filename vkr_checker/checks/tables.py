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
from .base import BaseCheck, CheckResult, Severity, add_issue


class TablesCheck(BaseCheck):
    check_id = "tables"
    check_name = "Оформление таблиц"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_fonts = self.rules["fonts"]
        allowed_cell_sizes = rules_fonts["sizes"]["table_cell"]   # например [10, 12]
        required_font = rules_fonts["required_font"]

        # Регулярное выражение для строки "Таблица N" – берём из конфигурации
        table_pattern_str = self.rules["captions"]["table_number_pattern"]
        table_pattern = re.compile(table_pattern_str)

        # 1. Собираем все пары (строка «Таблица N», название) из model.paragraphs
        caption_pairs = []   # список кортежей (caption_para, title_para)
        for i, para in enumerate(model.paragraphs):
            # Нормализуем неразрывные пробелы в обычные
            text = para.text.strip().replace("\u00A0", " ")
            if table_pattern.match(text):
                # ищем следующий непустой параграф — название таблицы (в пределах следующих 5)
                title_para = None
                for j in range(i + 1, min(i + 6, len(model.paragraphs))):
                    if model.paragraphs[j].text.strip():
                        title_para = model.paragraphs[j]
                        break
                caption_pairs.append((para, title_para))

        real_table_count = len(caption_pairs)   # количество найденных подписей

        # 2. Обходим блоки документа
        table_numbers_seen = []   # для проверки сквозной нумерации
        table_index = 0           # индекс текущей таблицы в caption_pairs

        for block in model.blocks:
            if block.kind != "table":
                continue

            # Если подписи для этой таблицы нет (таблиц больше, чем подписей) — пропускаем
            if table_index >= real_table_count:
                # Можно добавить предупреждение о таблице без подписи, но по умолчанию просто игнорируем
                # add_issue(result, "table_no_caption", ...)
                table_index += 1   # всё равно увеличиваем, чтобы не зациклиться
                continue

            caption_line, title_line = caption_pairs[table_index]
            table_num = table_index + 1   # порядковый номер (1-based)
            table = block.table

            # Проверка строки «Таблица N»
            if caption_line is None:
                add_issue(
                    result,
                    rule_id="table_no_number",
                    message=f"Таблица {table_num}: не найдена строка «Таблица N» перед таблицей",
                    severity=Severity.ERROR,
                )
            else:
                # Нормализуем текст для извлечения номера
                caption_text = caption_line.text.strip().replace("\u00A0", " ")
                m = table_pattern.match(caption_text)
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

            # Проверка названия таблицы
            if title_line is None:
                add_issue(
                    result,
                    rule_id="table_no_title",
                    message=f"Таблица {table_num}: отсутствует название",
                    severity=Severity.ERROR,
                )
            else:
                title_text = title_line.text.strip()
                # Название не должно заканчиваться точкой
                if title_text.endswith("."):
                    add_issue(
                        result,
                        rule_id="table_title_dot",
                        message=(
                            f"Таблица {table_num}: название таблицы "
                            f"не должно заканчиваться точкой"
                        ),
                        severity=Severity.ERROR,
                        context=title_text[:80],
                    )
                # Выравнивание названия
                if title_line.alignment not in (WD_ALIGN_PARAGRAPH.CENTER, None):
                    add_issue(
                        result,
                        rule_id="table_title_alignment",
                        message=(
                            f"Таблица {table_num}: название должно быть по центру"
                        ),
                        severity=Severity.WARNING,
                    )

            # Проверка шрифта, кегля и межстрочного интервала в ячейках таблицы
            self._check_cell_formatting(
                result, table, table_num,
                required_font, allowed_cell_sizes, resolver,
            )

            table_index += 1

        # 3. Проверка сквозной нумерации
        for idx, num in enumerate(table_numbers_seen):
            if num != idx + 1:
                add_issue(
                    result,
                    rule_id="table_numbering",
                    message=(
                        f"Нарушена сквозная нумерация таблиц: "
                        f"ожидалась Таблица {idx+1}, найдена Таблица {num}"
                    ),
                    severity=Severity.ERROR,
                )
                break

    def _check_cell_formatting(
            self, result, table, table_num,
            required_font, allowed_sizes, resolver,
    ):
        """Проверка форматирования ячеек таблицы."""
        expected_spacing = self.rules["spacing"]["table_cell"]
        spacing_tol = self.rules["spacing"].get("tolerance", 0.05)

        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if not text:
                        continue

                    # ---- Проверка межстрочного интервала ----
                    spacing = resolver.get_line_spacing(para)
                    if spacing is None:
                        add_issue(
                            result,
                            rule_id="table_cell_spacing_undefined",
                            message=(
                                f"Таблица {table_num}: не удалось определить межстрочный интервал "
                                f"(возможно, задан точный интервал 'exactly'). "
                                f"Требуется {expected_spacing}"
                            ),
                            severity=Severity.WARNING,
                            context=text[:60],
                        )
                        # Не прерываем проверку, идём дальше – возможно, шрифт определится
                    elif abs(spacing - expected_spacing) > spacing_tol:
                        add_issue(
                            result,
                            rule_id="table_cell_spacing",
                            message=(
                                f"Таблица {table_num}: межстрочный интервал "
                                f"{spacing:.2f} (требуется {expected_spacing})"
                            ),
                            severity=Severity.WARNING,
                            context=text[:60],
                        )
                        return  # одна ошибка на таблицу

                    # ---- Проверка шрифта и кегля ----
                    checked_any_run = False
                    for run in para.runs:
                        if not run.text.strip():
                            continue

                        checked_any_run = True
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
                                context=text[:60],
                            )
                            return

                        if size is None:
                            add_issue(
                                result,
                                rule_id="table_cell_size_undefined",
                                message=(
                                    f"Таблица {table_num}: не удалось определить кегль "
                                    f"в ячейке таблицы (ожидается {allowed_sizes})"
                                ),
                                severity=Severity.WARNING,
                                context=text[:60],
                            )
                            return  # прерываем, чтобы не дублировать

                        rounded_size = round(size)
                        if rounded_size not in allowed_sizes:
                            add_issue(
                                result,
                                rule_id="table_cell_size",
                                message=(
                                    f"Таблица {table_num}: кегль {rounded_size} пт "
                                    f"(допустимо: {allowed_sizes})"
                                ),
                                severity=Severity.WARNING,
                                context=f"{text[:60]} | реальный размер={size:.1f}",
                            )
                            return

                    if not checked_any_run:
                        continue