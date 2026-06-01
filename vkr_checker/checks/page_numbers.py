"""Проверка нумерации страниц."""
from __future__ import annotations

from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

from .base import BaseCheck, CheckResult, Severity, add_issue


class PageNumbersCheck(BaseCheck):
    check_id = "page_numbers"
    check_name = "Нумерация страниц"

    def _run(self, model, resolver, result: CheckResult) -> None:
        has_any_page_field = False

        for i, section in enumerate(model.doc.sections):
            footer = section.footer
            if not footer.is_linked_to_previous:
                page_field_found = False
                right_aligned = False

                for para in footer.paragraphs:
                    xml = para._p.xml
                    # PAGE field можно найти как fldChar + instrText или как простое поле
                    if "PAGE" in xml or "fldChar" in xml:
                        page_field_found = True
                        has_any_page_field = True
                    if page_field_found:
                        if para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                            right_aligned = True
                        elif para.alignment is None:
                            # Проверим через tab stops — иногда нумерация
                            # выровнена через правый таб
                            right_aligned = True  # допускаем

                if not page_field_found and i == 0:
                    add_issue(
                        result,
                        rule_id="no_page_number",
                        message=(
                            "Нумерация страниц не найдена в нижнем колонтитуле. "
                            "Номер должен стоять в правом нижнем углу"
                        ),
                        severity=Severity.ERROR,
                    )

                if page_field_found and not right_aligned:
                    add_issue(
                        result,
                        rule_id="page_number_alignment",
                        message=(
                            "Номер страницы должен быть выровнен по правому краю"
                        ),
                        severity=Severity.WARNING,
                        location_hint=f"Секция {i+1}",
                    )
