"""Проверка наличия подписей таблиц."""
from .base import BaseCheck, CheckResult, Severity, add_issue


class TablesCheck(BaseCheck):
    check_id = "tables"
    check_name = "Оформление таблиц"

    def _run(self, model, resolver, result: CheckResult) -> None:
        """Базовая проверка: наличие строки «Таблица N» перед каждой таблицей."""
        for block in model.blocks:
            if block.kind != "table":
                continue

            table_idx = block.index
            # Ищем предыдущий параграф (если он существует)
            if table_idx == 0:
                add_issue(
                    result,
                    "table_no_caption",
                    "Таблица без предшествующей подписи «Таблица N»",
                    severity=Severity.ERROR,
                    location_hint=f"~блок {table_idx + 1}",
                    context="Первая таблица в документе",
                )
                continue

            # Находим предыдущий параграф в blocks
            prev_block = None
            for b in model.blocks:
                if b.index == table_idx - 1:
                    prev_block = b
                    break

            if prev_block is None or prev_block.kind != "paragraph":
                add_issue(
                    result,
                    "table_no_caption",
                    "Таблица без предшествующей подписи «Таблица N»",
                    severity=Severity.ERROR,
                    location_hint=f"~блок {table_idx + 1}",
                    context="Нет параграфа перед таблицей",
                )
                continue

            prev_para = prev_block.paragraph
            prev_text = prev_para.text.strip()

            # Проверяем, соответствует ли текст шаблону «Таблица N»
            import re
            if not re.match(r"^Таблица\s+\d+\s*$", prev_text):
                add_issue(
                    result,
                    "table_no_caption",
                    f"Подпись таблицы не соответствует формату «Таблица N» (получено: «{prev_text}»)",
                    severity=Severity.ERROR,
                    location_hint=f"~блок {table_idx + 1}",
                    context=prev_text[:80],
                )
