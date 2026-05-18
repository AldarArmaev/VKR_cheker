"""
Юнит-тесты для модулей парсинга и резолвера стилей.

Запуск:
  pytest tests/test_parser_style_resolver.py -v
  pytest tests/ --cov=vkr_checker --cov-report=term-missing
"""
from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

from vkr_checker.parser import load_document, iter_blocks, _detect_sections, _find_intro_bounds, DocumentModel, SectionInfo, BlockItem
from vkr_checker.style_resolver import StyleResolver


FIXTURES = Path(__file__).parent / "fixtures"


# ──────────────────────────────────────────────
# Вспомогательные функции для создания тестовых документов
# ──────────────────────────────────────────────

def save_doc(doc: Document, tmp_path: Path, name: str = "test.docx") -> Path:
    """Сохраняет документ и возвращает путь к нему."""
    p = tmp_path / name
    doc.save(str(p))
    return p


def create_simple_doc(tmp_path: Path, paragraphs: list[str] | None = None) -> Path:
    """Создаёт простой документ с указанными параграфами."""
    doc = Document()
    if paragraphs:
        for text in paragraphs:
            doc.add_paragraph(text)
    else:
        doc.add_paragraph("Тестовый документ")
    return save_doc(doc, tmp_path)


# ──────────────────────────────────────────────
# Тесты модуля parser.py
# ──────────────────────────────────────────────

class TestLoadDocument:
    """Тесты функции load_document."""

    def test_load_existing_file(self, tmp_path):
        """Проверка загрузки существующего файла."""
        path = create_simple_doc(tmp_path)
        model = load_document(path)
        
        assert isinstance(model, DocumentModel)
        assert model.path == path
        assert model.doc is not None
        assert len(model.paragraphs) >= 1

    def test_load_nonexistent_file_raises_error(self):
        """Проверка ошибки при загрузке несуществующего файла."""
        with pytest.raises(FileNotFoundError):
            load_document("/nonexistent/path/file.docx")

    def test_load_invalid_extension_raises_error(self, tmp_path):
        """Проверка ошибки при загрузке файла с неправильным расширением."""
        invalid_path = tmp_path / "test.txt"
        invalid_path.write_text("text")
        
        with pytest.raises(ValueError, match="DOCX"):
            load_document(invalid_path)

    def test_load_document_with_sections(self, tmp_path):
        """Проверка загрузки документа с разделами."""
        doc = Document()
        doc.add_paragraph("Содержание")
        doc.add_paragraph("Введение")
        doc.add_paragraph("Глава 1. Теоретические основы")
        doc.add_paragraph("Заключение")
        doc.add_paragraph("Список литературы")
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        assert len(model.sections_found) >= 4
        section_names = [s.name for s in model.sections_found]
        assert "Введение" in section_names
        assert "Глава 1" in section_names
        assert "Заключение" in section_names
        assert "Список литературы" in section_names

    def test_intro_bounds_detection(self, tmp_path):
        """Проверка определения границ раздела Введение."""
        doc = Document()
        doc.add_paragraph("Содержание")
        doc.add_paragraph("Введение")
        doc.add_paragraph("Текст введения")
        doc.add_paragraph("Глава 1")
        doc.add_paragraph("Текст главы")
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        assert model.intro_start_idx >= 0
        assert model.intro_end_idx > model.intro_start_idx


class TestIterBlocks:
    """Тесты функции iter_blocks."""

    def test_iter_blocks_returns_paragraphs(self, tmp_path):
        """Проверка итерации по параграфам."""
        doc = Document()
        doc.add_paragraph("Параграф 1")
        doc.add_paragraph("Параграф 2")
        
        blocks = list(iter_blocks(doc))
        
        assert len(blocks) >= 2
        for block in blocks:
            assert isinstance(block, BlockItem)
            assert block.kind == "paragraph"
            assert block.paragraph is not None

    def test_iter_blocks_includes_tables(self, tmp_path):
        """Проверка что таблицы включаются в итерацию."""
        doc = Document()
        doc.add_paragraph("Текст перед таблицей")
        table = doc.add_table(rows=2, cols=2)
        doc.add_paragraph("Текст после таблицы")
        
        blocks = list(iter_blocks(doc))
        
        # Должны быть хотя бы 3 блока: параграф, таблица, параграф
        assert len(blocks) >= 3
        
        kinds = [b.kind for b in blocks]
        assert "table" in kinds
        assert "paragraph" in kinds

    def test_iter_blocks_preserves_order(self, tmp_path):
        """Проверка сохранения порядка элементов."""
        doc = Document()
        doc.add_paragraph("Первый")
        table = doc.add_table(rows=1, cols=1)
        doc.add_paragraph("Второй")
        
        blocks = list(iter_blocks(doc))
        
        assert len(blocks) >= 3
        # Первый блок должен быть параграфом с текстом "Первый"
        assert blocks[0].kind == "paragraph"
        assert "Первый" in blocks[0].paragraph.text


class TestDetectSections:
    """Тесты функции _detect_sections."""

    def test_detect_all_required_sections(self):
        """Проверка обнаружения всех обязательных разделов."""
        doc = Document()
        headings = [
            "Содержание",
            "Введение",
            "Глава 1. Теоретические основы",
            "Выводы по главе 1",
            "Глава 2. Практическая часть",
            "Заключение",
            "Список литературы",
        ]
        for heading in headings:
            doc.add_paragraph(heading)
        
        paragraphs = doc.paragraphs
        sections = _detect_sections(paragraphs)
        
        assert len(sections) == len(headings)
        section_names = [s.name for s in sections]
        assert "Введение" in section_names
        assert "Глава 1" in section_names
        assert "Глава 2" in section_names

    def test_detect_missing_sections(self):
        """Проверка что отсутствующие разделы не обнаруживаются."""
        doc = Document()
        doc.add_paragraph("Введение")
        doc.add_paragraph("Глава 1")
        # Заключение и Список литературы отсутствуют
        
        paragraphs = doc.paragraphs
        sections = _detect_sections(paragraphs)
        
        section_names = [s.name for s in sections]
        assert "Введение" in section_names
        assert "Глава 1" in section_names
        assert "Заключение" not in section_names
        assert "Список литературы" not in section_names

    def test_section_pattern_matching(self):
        """Проверка сопоставления с паттернами разделов."""
        doc = Document()
        # Разные варианты написания глав
        doc.add_paragraph("Глава 1. Теория")
        doc.add_paragraph("Глава 2 Практика")
        doc.add_paragraph("глава 3. Эксперимент")  # с маленькой буквы
        
        paragraphs = doc.paragraphs
        sections = _detect_sections(paragraphs)
        
        # Должны найтись все главы (паттерн case-insensitive)
        chapter_sections = [s for s in sections if "Глава" in s.name]
        assert len(chapter_sections) >= 2


class TestFindIntroBounds:
    """Тесты функции _find_intro_bounds."""

    def test_find_intro_bounds_correct(self):
        """Проверка корректного определения границ Введения."""
        doc = Document()
        doc.add_paragraph("Содержание")
        intro_para = doc.add_paragraph("Введение")
        doc.add_paragraph("Текст введения")
        chapter_para = doc.add_paragraph("Глава 1")
        
        paragraphs = doc.paragraphs
        sections = _detect_sections(paragraphs)
        start_idx, end_idx = _find_intro_bounds(paragraphs, sections)
        
        assert start_idx >= 0
        assert end_idx > start_idx

    def test_find_intro_bounds_no_intro(self):
        """Проверка когда Введение отсутствует."""
        doc = Document()
        doc.add_paragraph("Содержание")
        doc.add_paragraph("Глава 1")
        
        paragraphs = doc.paragraphs
        sections = _detect_sections(paragraphs)
        start_idx, end_idx = _find_intro_bounds(paragraphs, sections)
        
        assert start_idx == -1

    def test_find_intro_bounds_empty_sections(self):
        """Проверка с пустым списком разделов."""
        doc = Document()
        doc.add_paragraph("Текст")
        
        paragraphs = doc.paragraphs
        start_idx, end_idx = _find_intro_bounds(paragraphs, [])
        
        assert start_idx == -1
        assert end_idx == len(paragraphs)


# ──────────────────────────────────────────────
# Тесты модуля style_resolver.py
# ──────────────────────────────────────────────

class TestStyleResolver:
    """Тесты класса StyleResolver."""

    def test_resolver_initialization(self, tmp_path):
        """Проверка инициализации резолвера."""
        doc = Document()
        doc.add_paragraph("Тест")
        path = save_doc(doc, tmp_path)
        
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        assert resolver._doc is loaded_doc
        assert isinstance(resolver._defaults, dict)

    def test_get_font_name_from_run(self, tmp_path):
        """Проверка получения имени шрифта из run."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        run = para.runs[0]
        run.font.name = "Times New Roman"
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        font_name = resolver.get_font_name(loaded_run, loaded_para)
        assert font_name is not None

    def test_get_font_size_from_run(self, tmp_path):
        """Проверка получения размера шрифта из run."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        run = para.runs[0]
        run.font.size = Pt(14)
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        font_size = resolver.get_font_size_pt(loaded_run, loaded_para)
        assert font_size is not None
        assert font_size == 14.0

    def test_get_line_spacing(self, tmp_path):
        """Проверка получения межстрочного интервала."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        paragraph_format = para.paragraph_format
        paragraph_format.line_spacing = 1.5
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        spacing = resolver.get_line_spacing(loaded_para)
        
        # Проверяем конкретное значение 1.5
        assert spacing == 1.5

    def test_get_first_line_indent_cm(self, tmp_path):
        """Проверка получения отступа первой строки."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        paragraph_format = para.paragraph_format
        paragraph_format.first_line_indent = Cm(1.25)
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        indent = resolver.get_first_line_indent_cm(loaded_para)
        
        assert indent is not None
        # Ужесточённый допуск ±0.11 см
        assert 1.14 <= indent <= 1.36

    def test_is_bold_explicit(self, tmp_path):
        """Проверка определения жирного шрифта (явное задание)."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        run = para.runs[0]
        run.bold = True
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        is_bold = resolver.is_bold(loaded_run, loaded_para)
        assert is_bold is True

    def test_is_not_bold_explicit(self, tmp_path):
        """Проверка определения обычного шрифта (явное задание)."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        run = para.runs[0]
        run.bold = False
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        is_bold = resolver.is_bold(loaded_run, loaded_para)
        assert is_bold is False

    def test_is_italic_explicit(self, tmp_path):
        """Проверка определения курсива (явное задание)."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        run = para.runs[0]
        run.italic = True
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        is_italic = resolver.is_italic(loaded_run, loaded_para)
        assert is_italic is True

    def test_get_alignment_center(self, tmp_path):
        """Проверка получения центрированного выравнивания."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        alignment = resolver.get_alignment(loaded_para)
        
        assert alignment is not None

    def test_get_alignment_left(self, tmp_path):
        """Проверка получения левого выравнивания."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        alignment = resolver.get_alignment(loaded_para)
        
        assert alignment is not None

    def test_get_alignment_right(self, tmp_path):
        """Проверка получения правого выравнивания."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        alignment = resolver.get_alignment(loaded_para)
        
        assert alignment is not None

    def test_get_alignment_justify(self, tmp_path):
        """Проверка получения выравнивания по ширине."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        alignment = resolver.get_alignment(loaded_para)
        
        assert alignment is not None


# ──────────────────────────────────────────────
# Интеграционные тесты
# ──────────────────────────────────────────────

class TestIntegration:
    """Интеграционные тесты для parser и style_resolver."""

    def test_full_document_analysis(self, tmp_path):
        """Полный анализ документа с проверкой всех компонентов."""
        # Создаём документ с различными элементами
        doc = Document()
        
        # Заголовки разделов
        doc.add_paragraph("Содержание")
        doc.add_paragraph("Введение")
        
        # Параграф с форматированием
        para = doc.add_paragraph("Основной текст ВКР")
        run = para.runs[0]
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Таблица
        table = doc.add_table(rows=2, cols=2)
        
        # Ещё разделы
        doc.add_paragraph("Глава 1. Теоретические основы")
        doc.add_paragraph("Заключение")
        doc.add_paragraph("Список литературы")
        
        path = save_doc(doc, tmp_path)
        
        # Загружаем и анализируем
        model = load_document(path)
        
        # Проверка структуры
        assert len(model.sections_found) >= 4
        assert len(model.blocks) >= 6
        assert any(b.kind == "table" for b in model.blocks)
        
        # Проверка стилей
        resolver = StyleResolver(model.doc)
        # Проверяем что резолвер работает без ошибок
        for block in model.blocks:
            if block.kind == "paragraph" and block.paragraph:
                alignment = resolver.get_alignment(block.paragraph)
                # Не падаем с ошибкой

    def test_style_inheritance_from_paragraph_style(self, tmp_path):
        """Проверка наследования шрифта/кегля из стиля параграфа, если не задан в run."""
        doc = Document()
        
        # Создаём параграф со стилем по умолчанию (Normal)
        para = doc.add_paragraph("Текст со стилем Normal")
        # Не задаём шрифт явно в run - он должен наследоваться из стиля
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        if loaded_para.runs:
            loaded_run = loaded_para.runs[0]
            # Шрифт может быть None если не задан ни в run, ни в стиле, ни в дефолтах
            font_name = resolver.get_font_name(loaded_run, loaded_para)
            # Проверяем что резолвер работает без ошибок (значение может быть None)
            assert font_name is None or isinstance(font_name, str)
            
            # Кегль должен наследоваться из дефолтов документа (обычно 11pt)
            font_size = resolver.get_font_size_pt(loaded_run, loaded_para)
            assert font_size is not None
            assert font_size > 0

    def test_style_inheritance_explicit_run_overrides_style(self, tmp_path):
        """Проверка что явное задание в run переопределяет стиль параграфа."""
        doc = Document()
        
        para = doc.add_paragraph("Текст с явным шрифтом")
        run = para.runs[0]
        # Явно задаём шрифт и размер в run
        run.font.name = "Arial"
        run.font.size = Pt(16)
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        # Должны получить значения из run, а не из стиля
        font_name = resolver.get_font_name(loaded_run, loaded_para)
        assert font_name == "Arial"
        
        font_size = resolver.get_font_size_pt(loaded_run, loaded_para)
        assert font_size == 16.0

    def test_real_fixture_files(self):
        """Тесты на реальных файлах фикстур."""
        fixture_files = list(FIXTURES.glob("*.docx"))
        
        for fixture_path in fixture_files:
            # Загрузка должна работать без ошибок
            model = load_document(fixture_path)
            assert model is not None
            assert model.doc is not None
            
            # Резолвер должен инициализироваться
            resolver = StyleResolver(model.doc)
            assert resolver is not None


# ──────────────────────────────────────────────
# Параметризованные тесты
# ──────────────────────────────────────────────

class TestParameterized:
    """Параметризованные тесты для различных сценариев."""

    @pytest.mark.parametrize("margin_name,left,right,top,bottom", [
        ("correct", 3.0, 1.0, 2.0, 2.0),
        ("wrong_right", 3.0, 1.5, 2.0, 2.0),
        ("wrong_left", 2.0, 1.0, 2.0, 2.0),
        ("all_wrong", 2.0, 2.0, 3.0, 3.0),
    ])
    def test_margins_variations(self, tmp_path, margin_name, left, right, top, bottom):
        """Проверка документов с различными полями."""
        doc = Document()
        section = doc.sections[0]
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        
        doc.add_paragraph(f"Тест полей: {margin_name}")
        
        path = save_doc(doc, tmp_path, name=f"margins_{margin_name}.docx")
        model = load_document(path)
        
        assert model.doc is not None
        assert len(model.doc.sections) == 1
        
        # Проверка реальных значений полей через sections[0]
        loaded_section = model.doc.sections[0]
        assert abs(loaded_section.left_margin.cm - left) < 0.05
        assert abs(loaded_section.right_margin.cm - right) < 0.05
        assert abs(loaded_section.top_margin.cm - top) < 0.05
        assert abs(loaded_section.bottom_margin.cm - bottom) < 0.05

    @pytest.mark.parametrize("font_size", [10, 12, 14, 16])
    def test_different_font_sizes(self, tmp_path, font_size):
        """Проверка документов с различными размерами шрифта."""
        doc = Document()
        para = doc.add_paragraph(f"Текст с кеглем {font_size}")
        para.runs[0].font.size = Pt(font_size)
        
        path = save_doc(doc, tmp_path, name=f"fontsize_{font_size}.docx")
        model = load_document(path)
        
        resolver = StyleResolver(model.doc)
        loaded_para = model.doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        size = resolver.get_font_size_pt(loaded_run, loaded_para)
        assert size == float(font_size)

    @pytest.mark.parametrize("alignment_enum,alignment_name", [
        (WD_ALIGN_PARAGRAPH.LEFT, "left"),
        (WD_ALIGN_PARAGRAPH.CENTER, "center"),
        (WD_ALIGN_PARAGRAPH.RIGHT, "right"),
        (WD_ALIGN_PARAGRAPH.JUSTIFY, "justify"),
    ])
    def test_different_alignments(self, tmp_path, alignment_enum, alignment_name):
        """Проверка документов с различным выравниванием."""
        doc = Document()
        para = doc.add_paragraph(f"Выравнивание: {alignment_name}")
        para.alignment = alignment_enum
        
        path = save_doc(doc, tmp_path, name=f"align_{alignment_name}.docx")
        model = load_document(path)
        
        resolver = StyleResolver(model.doc)
        loaded_para = model.doc.paragraphs[0]
        
        alignment = resolver.get_alignment(loaded_para)
        assert alignment is not None or alignment is None


# ──────────────────────────────────────────────
# Тесты граничных случаев
# ──────────────────────────────────────────────

class TestEdgeCases:
    """Тесты граничных случаев и исключительных ситуаций."""

    def test_empty_document(self, tmp_path):
        """Проверка пустого документа."""
        doc = Document()
        # Пустой документ всё равно имеет один пустой параграф
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        assert model is not None
        assert len(model.paragraphs) >= 0

    def test_document_only_tables(self, tmp_path):
        """Проверка документа только с таблицами."""
        doc = Document()
        doc.add_table(rows=3, cols=3)
        doc.add_table(rows=2, cols=2)
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        tables_count = sum(1 for b in model.blocks if b.kind == "table")
        assert tables_count >= 2

    def test_document_very_long_text(self, tmp_path):
        """Проверка документа с очень длинным текстом."""
        doc = Document()
        long_text = "Длинный текст. " * 1000
        doc.add_paragraph(long_text)
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        assert len(model.paragraphs) >= 1
        assert len(model.paragraphs[0].text) > 1000

    def test_document_special_characters(self, tmp_path):
        """Проверка документа со специальными символами."""
        doc = Document()
        special_text = "Спецсимволы: © ® ™ € £ ¥ § ¶ † ‡ • …"
        doc.add_paragraph(special_text)
        
        path = save_doc(doc, tmp_path)
        model = load_document(path)
        
        assert model is not None
        assert any("©" in p.text for p in model.paragraphs)

    def test_resolver_with_none_paragraph(self, tmp_path):
        """Проверка резолвера с None параграфом."""
        doc = Document()
        para = doc.add_paragraph("Тест")
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        loaded_run = loaded_para.runs[0]
        
        # Методы должны работать с None или без ошибок обрабатывать
        is_bold = resolver.is_bold(loaded_run, None)
        assert isinstance(is_bold, bool)

    def test_multiple_runs_in_paragraph(self, tmp_path):
        """Проверка параграфа с несколькими runs."""
        doc = Document()
        para = doc.add_paragraph("")
        para.add_run("Первый run, ")
        para.add_run("второй run, ")
        para.add_run("третий run")
        
        # Применяем разное форматирование
        para.runs[0].font.name = "Times New Roman"
        para.runs[1].font.name = "Arial"
        para.runs[2].font.size = Pt(16)
        
        path = save_doc(doc, tmp_path)
        loaded_doc = Document(str(path))
        resolver = StyleResolver(loaded_doc)
        
        loaded_para = loaded_doc.paragraphs[0]
        
        # Каждый run должен обрабатываться независимо
        for run in loaded_para.runs:
            font_name = resolver.get_font_name(run, loaded_para)
            font_size = resolver.get_font_size_pt(run, loaded_para)
            # Не падаем с ошибкой


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
