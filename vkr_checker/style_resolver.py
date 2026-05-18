"""
Разворачивание цепочки наследования стилей в DOCX.

DOCX хранит форматирование на трёх уровнях:
  run.font  →  paragraph.style  →  parent style  →  Normal  →  document defaults

Если значение не задано явно на одном уровне, оно берётся с более высокого.
python-docx НЕ делает это автоматически, поэтому нужен StyleResolver.
"""
from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run


_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class StyleResolver:
    """Кэшированный резолвер стилей для одного документа."""

    def __init__(self, doc: Document):
        self._doc = doc
        self._defaults = self._load_document_defaults()

    # ──────────────────────────────────────────────
    # Публичные методы
    # ──────────────────────────────────────────────

    def get_font_name(self, run: Run, para: Paragraph) -> str | None:
        """Возвращает название шрифта с учётом всей цепочки наследования."""
        # 1. Прямое форматирование run
        if run.font.name:
            return run.font.name

        # 2. Цепочка стилей параграфа
        style = para.style
        while style:
            if style.font.name:
                return style.font.name
            style = style.base_style

        # 3. Дефолты документа
        return self._defaults.get("font_name")

    def get_font_size_pt(self, run: Run, para: Paragraph) -> float | None:
        """Возвращает размер шрифта в пунктах."""
        # 1. Прямое форматирование run
        if run.font.size:
            return run.font.size.pt

        # 2. Цепочка стилей
        style = para.style
        while style:
            if style.font.size:
                return style.font.size.pt
            style = style.base_style

        # 3. Дефолты документа
        sz_val = self._defaults.get("font_size_half_pt")
        return sz_val / 2 if sz_val else None

    def get_line_spacing(self, para: Paragraph) -> float | None:
        """
        Возвращает межстрочный интервал как множитель (1.0, 1.5, 2.0).
        None если не удалось определить.
        """
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return self._get_style_spacing(para)

        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            return self._get_style_spacing(para)

        line = spacing.get(qn("w:line"))
        rule = spacing.get(qn("w:lineRule"))

        if rule in ("auto", None) and line:
            # "auto" = proportional, 240 = single, 360 = 1.5x, 480 = double
            return round(int(line) / 240, 2)
        # "exact" / "atLeast" — абсолютное значение, не сравниваем как множитель
        return None

    def get_first_line_indent_cm(self, para: Paragraph) -> float:
        """Возвращает отступ первой строки в сантиметрах."""
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return 0.0

        ind = pPr.find(qn("w:ind"))
        if ind is None:
            return 0.0

        first_line = ind.get(qn("w:firstLine"))
        if first_line is None:
            return 0.0

        # Twips → cm: 1 twip = 1/567 cm (точнее: 1 inch = 1440 twips = 2.54 cm)
        return round(int(first_line) / 566.929, 3)

    def is_bold(self, run: Run, para: Paragraph | None = None) -> bool:
        """
        Проверяет жирность с учётом наследования.
        
        Алгоритм:
        1. Если run.bold задан явно (не None), используем его
        2. Иначе поднимаемся к Paragraph Style -> Base Style -> DocDefaults
        3. Проверяем style.font.bold в цепочке стилей
        """
        # 1. Прямое форматирование run
        if run.bold is not None:
            return run.bold
        
        # 2. Рекурсивный поиск в цепочке стилей параграфа
        if para is not None:
            style = para.style
            while style:
                # Проверяем font.bold стиля
                if hasattr(style, 'font') and style.font is not None:
                    if style.font.bold is not None:
                        return style.font.bold
                # Переходим к базовому стилю
                style = style.base_style
        
        # 3. Проверка дефолтов документа
        return self._defaults.get("bold", False)

    def is_italic(self, run: Run, para: Paragraph | None = None) -> bool:
        """
        Проверяет курсив с учётом наследования.
        Аналогично is_bold, поднимается по цепочке стилей.
        """
        # 1. Прямое форматирование run
        if run.italic is not None:
            return run.italic
        
        # 2. Рекурсивный поиск в цепочке стилей параграфа
        if para is not None:
            style = para.style
            while style:
                if hasattr(style, 'font') and style.font is not None:
                    if style.font.italic is not None:
                        return style.font.italic
                style = style.base_style
        
        # 3. Проверка дефолтов документа
        return self._defaults.get("italic", False)

    def get_alignment(self, para: Paragraph) -> str | None:
        """
        Возвращает выравнивание параграфа.
        Возможные значения: "left", "center", "right", "justify", None
        """
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            return self._get_style_alignment(para)
        
        jc = pPr.find(qn("w:jc"))
        if jc is None:
            return self._get_style_alignment(para)
        
        val = jc.get(qn("w:val"))
        if val:
            # Нормализация значений
            mapping = {
                "left": "left",
                "start": "left",
                "center": "center",
                "right": "right",
                "end": "right",
                "both": "justify",
                "justify": "justify",
                "distribute": "justify",
            }
            return mapping.get(val.lower(), val)
        
        return self._get_style_alignment(para)

    # ──────────────────────────────────────────────
    # Приватные методы
    # ──────────────────────────────────────────────

    def _load_document_defaults(self) -> dict:
        """Читает w:docDefaults из styles.xml."""
        defaults = {}
        doc_defaults = self._doc.styles.element.find(
            f".//{{{_NS}}}docDefaults"
        )
        if doc_defaults is None:
            return defaults

        rpr_default = doc_defaults.find(f".//{{{_NS}}}rPrDefault//{{{_NS}}}rPr")
        if rpr_default is None:
            return defaults

        # Шрифт
        r_fonts = rpr_default.find(f"{{{_NS}}}rFonts")
        if r_fonts is not None:
            defaults["font_name"] = (
                r_fonts.get(f"{{{_NS}}}ascii")
                or r_fonts.get(f"{{{_NS}}}hAnsi")
            )

        # Размер (в half-points)
        sz = rpr_default.find(f"{{{_NS}}}sz")
        if sz is not None:
            val = sz.get(f"{{{_NS}}}val")
            if val:
                defaults["font_size_half_pt"] = int(val)

        # Жирность
        b = rpr_default.find(f"{{{_NS}}}b")
        if b is not None:
            val = b.get(f"{{{_NS}}}val")
            defaults["bold"] = val != "false"

        # Курсив
        i = rpr_default.find(f"{{{_NS}}}i")
        if i is not None:
            val = i.get(f"{{{_NS}}}val")
            defaults["italic"] = val != "false"

        return defaults

    def _get_style_spacing(self, para: Paragraph) -> float | None:
        """Читает spacing из цепочки стилей."""
        style = para.style
        while style:
            if style._element is not None:
                pPr = style._element.find(f"{{{_NS}}}pPr")
                if pPr is not None:
                    spacing = pPr.find(f"{{{_NS}}}spacing")
                    if spacing is not None:
                        line = spacing.get(f"{{{_NS}}}line")
                        rule = spacing.get(f"{{{_NS}}}lineRule")
                        if rule in ("auto", None) and line:
                            return round(int(line) / 240, 2)
            style = style.base_style
        return None

    def _get_style_alignment(self, para: Paragraph) -> str | None:
        """Читает выравнивание из цепочки стилей."""
        style = para.style
        while style:
            if style._element is not None:
                pPr = style._element.find(f"{{{_NS}}}pPr")
                if pPr is not None:
                    jc = pPr.find(f"{{{_NS}}}jc")
                    if jc is not None:
                        val = jc.get(f"{{{_NS}}}val")
                        if val:
                            mapping = {
                                "left": "left",
                                "start": "left",
                                "center": "center",
                                "right": "right",
                                "end": "right",
                                "both": "justify",
                                "justify": "justify",
                            }
                            return mapping.get(val.lower(), val)
            style = style.base_style
        return None
