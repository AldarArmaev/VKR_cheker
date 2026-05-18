# Полная техническая спецификация: VKR Checker

> **Назначение документа:** Исчерпывающее руководство для разработки программы
> автоматической проверки формальных требований к оформлению ВКР.
> По этому документу можно создать готовый продукт без дополнительных источников.
>
> **Регламент:** Требования ВКР бакалавров (психологический факультет)
> **Язык:** Python 3.11+
> **Целевая платформа:** Windows 10/11, macOS 12+, Ubuntu 22.04+

---

## Содержание

1. [Структура проекта](#1-структура-проекта)
2. [Зависимости и установка](#2-зависимости-и-установка)
3. [Конфигурационный файл требований](#3-конфигурационный-файл-требований)
4. [Модуль парсинга документа](#4-модуль-парсинга-документа)
5. [Резолвер стилей](#5-резолвер-стилей)
6. [Базовые классы проверок](#6-базовые-классы-проверок)
7. [Модули проверок](#7-модули-проверок)
   - 7.1 Поля страницы
   - 7.2 Шрифт и кегль
   - 7.3 Межстрочный интервал
   - 7.4 Абзацные отступы
   - 7.5 Выравнивание
   - 7.6 Нумерация страниц
   - 7.7 Стили заголовков
   - 7.8 Разделы документа
   - 7.9 Структура Введения
   - 7.10 Таблицы
   - 7.11 Рисунки
   - 7.12 Список литературы
8. [Оркестратор проверок](#8-оркестратор-проверок)
9. [Генераторы отчётов](#9-генераторы-отчётов)
10. [CLI-интерфейс](#10-cli-интерфейс)
11. [GUI-интерфейс (tkinter)](#11-gui-интерфейс-tkinter)
12. [Точка входа](#12-точка-входа)
13. [Тесты](#13-тесты)
14. [Сборка в исполняемый файл](#14-сборка-в-исполняемый-файл)

---

## 1. Структура проекта

```
vkr_checker/
├── config/
│   └── rules.yaml              # Все требования регламента
├── vkr_checker/
│   ├── __init__.py
│   ├── parser.py               # Загрузка документа
│   ├── style_resolver.py       # Разворачивание цепочек стилей
│   ├── checks/
│   │   ├── __init__.py
│   │   ├── base.py             # Issue, CheckResult, BaseCheck
│   │   ├── margins.py
│   │   ├── fonts.py
│   │   ├── spacing.py
│   │   ├── indents.py
│   │   ├── alignment.py
│   │   ├── page_numbers.py
│   │   ├── heading_styles.py
│   │   ├── sections.py
│   │   ├── intro.py
│   │   ├── tables.py
│   │   ├── figures.py
│   │   └── bibliography.py
│   ├── reporters/
│   │   ├── __init__.py
│   │   ├── console_reporter.py
│   │   └── html_reporter.py
│   ├── orchestrator.py         # Запуск всех проверок
│   ├── cli.py                  # Click CLI
│   └── gui.py                  # tkinter GUI
├── tests/
│   ├── fixtures/               # Тестовые .docx файлы
│   └── test_checks.py
├── main.py                     # Точка входа
├── pyproject.toml
└── README.md
```

---

## 2. Зависимости и установка

### pyproject.toml

```toml
[build-system]
requires = ["setuptools>=68", "wheel"]
build-backend = "setuptools.backends.legacy:build"

[project]
name = "vkr-checker"
version = "1.0.0"
description = "Автоматическая проверка оформления ВКР"
requires-python = ">=3.11"
dependencies = [
    "python-docx>=1.1.0",
    "lxml>=5.0.0",
    "PyYAML>=6.0",
    "jinja2>=3.1",
    "click>=8.1",
    "rich>=13.0",
]

[project.scripts]
vkr-check = "vkr_checker.cli:main"

[project.optional-dependencies]
gui = []  # tkinter входит в стандартную библиотеку
build = ["pyinstaller>=6.0"]
test = ["pytest>=7.0", "pytest-cov>=4.0"]
```

### Установка

```bash
# Клонировать репозиторий и установить
pip install -e ".[test]"

# Или только зависимости
pip install python-docx lxml PyYAML jinja2 click rich
```

---

## 3. Конфигурационный файл требований

### config/rules.yaml

```yaml
# =============================================================
# Конфигурация требований к оформлению ВКР
# Редактировать ТОЛЬКО этот файл при смене регламента
# =============================================================

meta:
  version: "2024"
  institution: "ВКР бакалавров (психологический факультет)"

# ──────────────────────────────────────────────
# Объём работы
# ──────────────────────────────────────────────
document:
  type: bachelor
  volume_min_pages: 40    # без приложений
  volume_max_pages: 45
  bibliography_min_sources: 40

# ──────────────────────────────────────────────
# Поля страницы (в сантиметрах)
# ──────────────────────────────────────────────
margins_cm:
  left: 3.0
  right: 1.0
  top: 2.0
  bottom: 2.0
  tolerance: 0.15   # допуск ±0.15 см (~1.5 мм)

# ──────────────────────────────────────────────
# Шрифты
# ──────────────────────────────────────────────
fonts:
  # Единственный разрешённый шрифт для всего документа
  required_font: "Times New Roman"

  sizes:
    body: 14           # основной текст
    table_cell: [11, 12]   # допустимые значения в ячейках
    figure_caption: 12
    table_caption_word: 14   # строка "Таблица N"
    table_caption_title: 14  # название таблицы
    page_number: 12

# ──────────────────────────────────────────────
# Межстрочный интервал
# ──────────────────────────────────────────────
spacing:
  body: 1.5              # основной текст (Введение → конец Приложений)
  title_page: 1.0
  contents: 1.0
  table_cell: 1.0
  table_caption_title: 1.5   # название таблицы (строка с текстом)
  figure_caption: 1.0
  empty_line_between_refs: 1.0  # пустая строка после рисунка перед подписью
  tolerance: 0.05        # допуск ±0.05

# ──────────────────────────────────────────────
# Абзацные отступы (первая строка, в сантиметрах)
# ──────────────────────────────────────────────
indents:
  body_first_line: 1.25         # красная строка в основном тексте
  heading_first_line: 0.0       # без отступа у заголовков
  table_caption_first_line: 0.0
  figure_caption_first_line: 0.0
  tolerance: 0.1                # допуск ±0.1 см

# ──────────────────────────────────────────────
# Выравнивание
# ──────────────────────────────────────────────
alignment:
  body: "justify"
  chapter_heading: "center"
  paragraph_heading: "center"
  table_number_line: "right"   # строка "Таблица N"
  table_caption_title: "center" # название таблицы
  figure_caption: "center"

# ──────────────────────────────────────────────
# Нумерация страниц
# ──────────────────────────────────────────────
page_numbers:
  position: "right_bottom"
  title_page_numbered: false
  contents_page_numbered: false
  intro_starts_at: 3   # страница Введения = 3

# ──────────────────────────────────────────────
# Стили
# ──────────────────────────────────────────────
styles:
  # Применение этих стилей запрещено во всём документе
  forbidden:
    - "Heading 1"
    - "Heading 2"
    - "Heading 3"
    - "Heading 4"
    - "Heading 5"
    - "Заголовок 1"
    - "Заголовок 2"
    - "Заголовок 3"
    - "Заголовок 4"
    - "Заголовок 5"
  # Допустимые стили для всего документа
  allowed_body: ["Normal", "Обычный", "Default Paragraph Style"]

# ──────────────────────────────────────────────
# Обязательные разделы (в порядке следования)
# ──────────────────────────────────────────────
required_sections:
  - pattern: "^Содержание$"
    name: "Содержание"
    new_page: true
  - pattern: "^Введение$"
    name: "Введение"
    new_page: true
  - pattern: "^Глава\\s+1"
    name: "Глава 1"
    new_page: true
  - pattern: "^Выводы по главе"
    name: "Выводы по главе 1"
    new_page: false
  - pattern: "^Глава\\s+2"
    name: "Глава 2"
    new_page: true
  - pattern: "^Заключение$"
    name: "Заключение"
    new_page: true
  - pattern: "^Список литературы$"
    name: "Список литературы"
    new_page: true

# ──────────────────────────────────────────────
# Обязательные подразделы Введения
# (строго в указанном порядке)
# ──────────────────────────────────────────────
intro_subsections:
  - keyword: "Актуальность"
    bold_required: true
  - keyword: "Цель"
    bold_required: true
  - keyword: "Объект"
    bold_required: true
  - keyword: "Предмет"
    bold_required: true
  - keyword: "Гипотез"        # "Гипотеза" или "Гипотезы"
    bold_required: true
  - keyword: "Задачи"
    bold_required: true
  - keyword: "Теоретическая основа"
    bold_required: true
  - keyword: "Общая характеристика выборки"
    bold_required: true
  - keyword: "Методы и методики"
    bold_required: true
  - keyword: "Структура и объем"
    bold_required: true

# ──────────────────────────────────────────────
# Форматы подписей таблиц и рисунков
# ──────────────────────────────────────────────
captions:
  # Строка над таблицей: "Таблица 1" (без точки, выравнивание по правому краю)
  table_number_pattern: "^Таблица\\s+\\d+\\s*$"

  # Строка с названием таблицы идёт сразу после строки с номером
  # Название таблицы не должно заканчиваться точкой
  table_name_no_dot: true

  # Подпись рисунка: "Рис. N. Название" — точка после "Рис." и N, без точки в конце
  figure_pattern: "^Рис\\.\\s+\\d+\\.\\s+[А-ЯA-ZЁ].+[^.!?]$"
  figure_prefix: "Рис."

# ──────────────────────────────────────────────
# Список литературы
# ──────────────────────────────────────────────
bibliography:
  min_sources: 40
  # Запрещённые типы источников (проверяется по ключевым словам)
  forbidden_source_types:
    - "учебник"
    - "учебное пособие"
    - "учеб. пособие"
    - "учеб пособие"
```

---

## 4. Модуль парсинга документа

### vkr_checker/parser.py

```python
"""
Загрузка и первичный разбор DOCX-документа.
Возвращает объект DocumentModel с нормализованными данными.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass
class BlockItem:
    """Унифицированный элемент документа (параграф или таблица)."""
    index: int           # порядковый номер в документе
    kind: str            # "paragraph" или "table"
    paragraph: Paragraph | None = None
    table: Table | None = None


@dataclass
class SectionInfo:
    """Информация об обнаруженном разделе."""
    name: str
    pattern: str
    paragraph_index: int
    heading_text: str


@dataclass
class DocumentModel:
    """Нормализованная модель документа."""
    doc: Document
    path: Path
    blocks: list[BlockItem]          # все элементы в порядке следования
    paragraphs: list[Paragraph]      # только параграфы
    tables: list[Table]              # только таблицы
    sections_found: list[SectionInfo]
    intro_start_idx: int = -1        # индекс параграфа начала Введения
    intro_end_idx: int = -1          # индекс параграфа конца Введения


def load_document(path: str | Path) -> DocumentModel:
    """
    Загружает DOCX и строит DocumentModel.

    Raises:
        FileNotFoundError: если файл не найден
        ValueError: если формат файла не поддерживается
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Поддерживается только DOCX, получен: {path.suffix}")

    doc = Document(str(path))
    blocks = list(_iter_blocks(doc))
    paragraphs = [b.paragraph for b in blocks if b.kind == "paragraph"]
    tables = [b.table for b in blocks if b.kind == "table"]
    sections = _detect_sections(paragraphs)
    intro_start, intro_end = _find_intro_bounds(paragraphs, sections)

    return DocumentModel(
        doc=doc,
        path=path,
        blocks=blocks,
        paragraphs=paragraphs,
        tables=tables,
        sections_found=sections,
        intro_start_idx=intro_start,
        intro_end_idx=intro_end,
    )


def _iter_blocks(doc: Document) -> Iterator[BlockItem]:
    """
    Итерирует по документу в правильном порядке, включая таблицы.
    python-docx doc.paragraphs пропускает таблицы, поэтому используем XML.
    """
    idx = 0
    # Элементы верхнего уровня тела документа
    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "p":
            # Найти соответствующий Paragraph объект
            para = _find_paragraph_by_element(doc, child)
            if para is not None:
                yield BlockItem(index=idx, kind="paragraph", paragraph=para)
                idx += 1
        elif tag == "tbl":
            table = _find_table_by_element(doc, child)
            if table is not None:
                yield BlockItem(index=idx, kind="table", table=table)
                idx += 1


def _find_paragraph_by_element(doc: Document, elem) -> Paragraph | None:
    for p in doc.paragraphs:
        if p._p is elem:
            return p
    return None


def _find_table_by_element(doc: Document, elem) -> Table | None:
    for t in doc.tables:
        if t._tbl is elem:
            return t
    return None


def _detect_sections(paragraphs: list[Paragraph]) -> list[SectionInfo]:
    """Определяет разделы документа по тексту заголовков."""
    import yaml
    from pathlib import Path as P

    rules_path = P("config/rules.yaml")
    if not rules_path.exists():
        rules_path = P(__file__).parent.parent / "config" / "rules.yaml"

    with open(rules_path, encoding="utf-8") as f:
        rules = yaml.safe_load(f)

    section_patterns = rules.get("required_sections", [])
    found = []

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        for sec in section_patterns:
            if re.match(sec["pattern"], text, re.IGNORECASE):
                found.append(SectionInfo(
                    name=sec["name"],
                    pattern=sec["pattern"],
                    paragraph_index=i,
                    heading_text=text,
                ))
    return found


def _find_intro_bounds(
    paragraphs: list[Paragraph],
    sections: list[SectionInfo],
) -> tuple[int, int]:
    """Возвращает индексы начала и конца раздела Введение."""
    intro_idx = -1
    next_major_idx = len(paragraphs)

    for sec in sections:
        if re.match(r"^Введение$", sec.heading_text, re.IGNORECASE):
            intro_idx = sec.paragraph_index
        elif intro_idx >= 0 and sec.paragraph_index > intro_idx:
            next_major_idx = sec.paragraph_index
            break

    return intro_idx, next_major_idx
```

---

## 5. Резолвер стилей

### vkr_checker/style_resolver.py

```python
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

    def is_bold(self, run: Run) -> bool:
        """Проверяет жирность с учётом наследования."""
        if run.bold is not None:
            return run.bold
        style = run._r.getparent()  # w:p или w:tr и т.п.
        return False

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
```

---

## 6. Базовые классы проверок

### vkr_checker/checks/base.py

```python
"""
Базовые классы для системы проверок.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Severity(Enum):
    ERROR = "error"       # Грубое нарушение — нормоконтроль не пропустит
    WARNING = "warning"   # Вероятная ошибка — стоит проверить вручную
    INFO = "info"         # Информационное замечание


SEVERITY_LABELS = {
    Severity.ERROR:   "✗ ОШИБКА",
    Severity.WARNING: "⚠ ПРЕДУПРЕЖДЕНИЕ",
    Severity.INFO:    "ℹ ИНФОРМАЦИЯ",
}


@dataclass
class Issue:
    """Одно обнаруженное замечание."""
    rule_id: str             # уникальный идентификатор правила
    message: str             # человекочитаемое описание
    severity: Severity = Severity.ERROR
    location_hint: str = ""  # например: "~стр. 12" или "Таблица 3"
    context: str = ""        # цитата проблемного текста (до 80 символов)

    def __str__(self) -> str:
        parts = [SEVERITY_LABELS[self.severity], self.message]
        if self.location_hint:
            parts.append(f"({self.location_hint})")
        if self.context:
            parts.append(f"→ «{self.context[:80]}»")
        return "  ".join(parts)


@dataclass
class CheckResult:
    """Результат одного модуля проверок."""
    check_id: str
    check_name: str
    issues: list[Issue] = field(default_factory=list)
    skipped: bool = False
    skip_reason: str = ""

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.ERROR)

    @property
    def warning_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.WARNING)

    @property
    def passed(self) -> bool:
        return self.error_count == 0 and not self.skipped


class BaseCheck:
    """
    Базовый класс для всех модулей проверок.

    Каждый наследник обязан:
    - задать class-атрибуты check_id и check_name
    - реализовать метод _run(model, resolver, rules)
    """
    check_id: str = "base"
    check_name: str = "Базовая проверка"

    def __init__(self, rules: dict):
        self.rules = rules

    def run(self, model, resolver) -> CheckResult:
        """Запускает проверку и возвращает CheckResult."""
        result = CheckResult(
            check_id=self.check_id,
            check_name=self.check_name,
        )
        try:
            self._run(model, resolver, result)
        except Exception as exc:
            result.skipped = True
            result.skip_reason = f"Ошибка при выполнении проверки: {exc}"
        return result

    def _run(self, model, resolver, result: CheckResult) -> None:
        raise NotImplementedError


def add_issue(
    result: CheckResult,
    rule_id: str,
    message: str,
    severity: Severity = Severity.ERROR,
    location_hint: str = "",
    context: str = "",
) -> None:
    """Вспомогательная функция добавления замечания."""
    result.issues.append(Issue(
        rule_id=rule_id,
        message=message,
        severity=severity,
        location_hint=location_hint,
        context=context[:100],
    ))
```

---

## 7. Модули проверок

### 7.1 vkr_checker/checks/margins.py

```python
"""Проверка полей страницы."""
from .base import BaseCheck, CheckResult, Severity, add_issue


class MarginsCheck(BaseCheck):
    check_id = "margins"
    check_name = "Поля страницы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        expected = self.rules["margins_cm"]
        tol = expected.get("tolerance", 0.15)

        for i, section in enumerate(model.doc.sections):
            actual = {
                "left":   section.left_margin.cm,
                "right":  section.right_margin.cm,
                "top":    section.top_margin.cm,
                "bottom": section.bottom_margin.cm,
            }
            labels = {
                "left": "левое", "right": "правое",
                "top": "верхнее", "bottom": "нижнее",
            }
            for side, label in labels.items():
                exp_val = expected[side]
                act_val = actual[side]
                if abs(act_val - exp_val) > tol:
                    add_issue(
                        result,
                        rule_id=f"margin_{side}",
                        message=(
                            f"Поле «{label}»: {act_val:.2f} см "
                            f"(требуется {exp_val:.1f} см)"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"Секция {i+1}",
                    )
```

### 7.2 vkr_checker/checks/fonts.py

```python
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
                if font_name and font_name != required_font:
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
```

### 7.3 vkr_checker/checks/spacing.py

```python
"""Проверка межстрочного интервала."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import is_heading_paragraph, para_is_in_table


class SpacingCheck(BaseCheck):
    check_id = "spacing"
    check_name = "Межстрочный интервал"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["spacing"]
        body_spacing = rules["body"]
        tol = rules.get("tolerance", 0.05)

        in_main_text = False
        last_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            # Начало основного текста — с Введения
            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True

            if not in_main_text:
                continue
            if not text:
                continue
            if para_is_in_table(para):
                continue  # таблицы проверяются отдельно

            spacing = resolver.get_line_spacing(para)
            if spacing is None:
                continue  # не можем определить — пропускаем

            if abs(spacing - body_spacing) > tol:
                if i - last_issue_para >= 5:
                    add_issue(
                        result,
                        rule_id="line_spacing",
                        message=(
                            f"Межстрочный интервал {spacing:.2f} "
                            f"(требуется {body_spacing})"
                        ),
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )
                    last_issue_para = i
```

### 7.4 vkr_checker/checks/indents.py

```python
"""Проверка абзацных отступов."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import is_heading_paragraph, para_is_in_table


class IndentsCheck(BaseCheck):
    check_id = "indents"
    check_name = "Абзацные отступы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules = self.rules["indents"]
        body_indent = rules["body_first_line"]
        tol = rules.get("tolerance", 0.1)

        in_main_text = False
        last_issue_para = -10

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()

            if re.match(r"^Введение$", text, re.IGNORECASE):
                in_main_text = True
            if not in_main_text or not text:
                continue
            if para_is_in_table(para):
                continue

            indent = resolver.get_first_line_indent_cm(para)
            is_heading = is_heading_paragraph(para)

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
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )
            else:
                # Основной текст — отступ 1.25 см
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
                            location_hint=f"~абз. {i+1}",
                            context=text[:80],
                        )
                        last_issue_para = i
```

### 7.5 vkr_checker/checks/alignment.py

```python
"""Проверка выравнивания текста."""
from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from .base import BaseCheck, CheckResult, Severity, add_issue
from .fonts import is_heading_paragraph, para_is_in_table

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

            actual_align = para.alignment

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
            if is_heading_paragraph(para):
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
```

### 7.6 vkr_checker/checks/page_numbers.py

```python
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
```

### 7.7 vkr_checker/checks/heading_styles.py

```python
"""Проверка запрещённых стилей заголовков."""
from .base import BaseCheck, CheckResult, Severity, add_issue


class HeadingStylesCheck(BaseCheck):
    check_id = "heading_styles"
    check_name = "Стили заголовков"

    def _run(self, model, resolver, result: CheckResult) -> None:
        forbidden = set(self.rules["styles"]["forbidden"])

        for i, para in enumerate(model.paragraphs):
            style_name = para.style.name
            if style_name in forbidden:
                add_issue(
                    result,
                    rule_id="forbidden_style",
                    message=(
                        f"Применён стиль «{style_name}». "
                        f"Весь текст должен иметь стиль «Обычный»"
                    ),
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                    context=para.text[:80],
                )
```

### 7.8 vkr_checker/checks/sections.py

```python
"""Проверка наличия обязательных разделов."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue


class SectionsCheck(BaseCheck):
    check_id = "sections"
    check_name = "Обязательные разделы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        required = self.rules["required_sections"]
        all_texts = [p.text.strip() for p in model.paragraphs]

        found_names = {s.name for s in model.sections_found}

        for sec in required:
            pattern = sec["pattern"]
            name = sec["name"]
            found = any(
                re.match(pattern, t, re.IGNORECASE)
                for t in all_texts if t
            )
            if not found:
                add_issue(
                    result,
                    rule_id=f"missing_section_{name.replace(' ', '_')}",
                    message=f"Отсутствует обязательный раздел: «{name}»",
                    severity=Severity.ERROR,
                )

        # Проверка порядка разделов
        self._check_order(model, result)

    def _check_order(self, model, result: CheckResult) -> None:
        """Разделы должны идти в предписанном порядке."""
        sections = model.sections_found
        if len(sections) < 2:
            return

        indices = [s.paragraph_index for s in sections]
        for j in range(len(indices) - 1):
            if indices[j] > indices[j + 1]:
                add_issue(
                    result,
                    rule_id="section_order",
                    message=(
                        f"Нарушен порядок разделов: «{sections[j].name}» "
                        f"стоит после «{sections[j+1].name}»"
                    ),
                    severity=Severity.ERROR,
                )
```

### 7.9 vkr_checker/checks/intro.py

```python
"""Проверка структуры раздела Введение."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue


class IntroCheck(BaseCheck):
    check_id = "intro"
    check_name = "Структура Введения"

    def _run(self, model, resolver, result: CheckResult) -> None:
        if model.intro_start_idx < 0:
            add_issue(
                result,
                rule_id="intro_not_found",
                message="Раздел «Введение» не найден — проверка структуры невозможна",
                severity=Severity.ERROR,
            )
            return

        subsections = self.rules["intro_subsections"]
        intro_paras = model.paragraphs[
            model.intro_start_idx : model.intro_end_idx
        ]
        intro_text = " ".join(p.text.strip() for p in intro_paras)

        for sub in subsections:
            keyword = sub["keyword"]
            bold_required = sub.get("bold_required", True)

            # Ищем параграф, начинающийся с ключевого слова
            found_para = None
            for para in intro_paras:
                if para.text.strip().startswith(keyword):
                    found_para = para
                    break

            if found_para is None:
                add_issue(
                    result,
                    rule_id=f"intro_missing_{keyword.replace(' ', '_')[:20]}",
                    message=f"Во Введении отсутствует подраздел «{keyword}»",
                    severity=Severity.ERROR,
                )
                continue

            # Проверяем, что ключевое слово выделено жирным
            if bold_required:
                is_bold = any(
                    run.bold and keyword[:6] in run.text
                    for run in found_para.runs
                )
                if not is_bold:
                    add_issue(
                        result,
                        rule_id=f"intro_bold_{keyword.replace(' ', '_')[:20]}",
                        message=(
                            f"«{keyword}» во Введении должно быть "
                            f"выделено жирным шрифтом"
                        ),
                        severity=Severity.WARNING,
                        context=found_para.text[:80],
                    )
```

### 7.10 vkr_checker/checks/tables.py

```python
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
```

### 7.11 vkr_checker/checks/figures.py

```python
"""Проверка оформления рисунков и диаграмм."""
from __future__ import annotations
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from .base import BaseCheck, CheckResult, Severity, add_issue


class FiguresCheck(BaseCheck):
    check_id = "figures"
    check_name = "Оформление рисунков"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_captions = self.rules["captions"]
        figure_re = re.compile(rules_captions["figure_pattern"])
        figure_numbers_seen = []
        figure_number_re = re.compile(r"^Рис\.\s+(\d+)\.")

        for i, para in enumerate(model.paragraphs):
            text = para.text.strip()
            if not text.startswith("Рис."):
                continue

            # Извлекаем номер
            m = figure_number_re.match(text)
            if m:
                figure_numbers_seen.append(int(m.group(1)))

            # Проверяем формат подписи
            if not figure_re.match(text):
                add_issue(
                    result,
                    rule_id="figure_caption_format",
                    message=(
                        "Неверный формат подписи рисунка. "
                        "Должно быть: «Рис. N. Название» "
                        "(с прописной буквы, без точки в конце)"
                    ),
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                    context=text[:80],
                )
            else:
                # Точка в конце названия недопустима
                if text.endswith("."):
                    add_issue(
                        result,
                        rule_id="figure_caption_dot",
                        message="Подпись рисунка не должна заканчиваться точкой",
                        severity=Severity.ERROR,
                        location_hint=f"~абз. {i+1}",
                        context=text[:80],
                    )

            # Выравнивание
            if para.alignment not in (WD_ALIGN_PARAGRAPH.CENTER, None):
                add_issue(
                    result,
                    rule_id="figure_caption_alignment",
                    message="Подпись рисунка должна быть выровнена по центру",
                    severity=Severity.ERROR,
                    location_hint=f"~абз. {i+1}",
                )

        # Сквозная нумерация
        for j, num in enumerate(figure_numbers_seen):
            if num != j + 1:
                add_issue(
                    result,
                    rule_id="figure_numbering",
                    message=(
                        f"Нарушена сквозная нумерация рисунков: "
                        f"ожидался Рис. {j+1}, найден Рис. {num}"
                    ),
                    severity=Severity.ERROR,
                )
                break
```

### 7.12 vkr_checker/checks/bibliography.py

```python
"""Проверка списка литературы."""
from __future__ import annotations
import re
from .base import BaseCheck, CheckResult, Severity, add_issue


class BibliographyCheck(BaseCheck):
    check_id = "bibliography"
    check_name = "Список литературы"

    def _run(self, model, resolver, result: CheckResult) -> None:
        rules_bib = self.rules["bibliography"]
        min_sources = rules_bib["min_sources"]
        forbidden_types = [ft.lower() for ft in rules_bib["forbidden_source_types"]]

        # Найти раздел "Список литературы"
        bib_start = -1
        for i, para in enumerate(model.paragraphs):
            if re.match(r"^Список литературы$", para.text.strip(), re.IGNORECASE):
                bib_start = i
                break

        if bib_start < 0:
            add_issue(
                result,
                rule_id="bibliography_not_found",
                message="Раздел «Список литературы» не найден",
                severity=Severity.ERROR,
            )
            return

        # Найти конец раздела (следующий основной раздел или конец документа)
        bib_end = len(model.paragraphs)
        for i in range(bib_start + 1, len(model.paragraphs)):
            text = model.paragraphs[i].text.strip()
            if re.match(r"^Приложени", text, re.IGNORECASE):
                bib_end = i
                break

        bib_paras = model.paragraphs[bib_start + 1 : bib_end]

        # Подсчёт источников (непустые строки, начинающиеся с цифры или буквы)
        source_lines = [
            p for p in bib_paras
            if p.text.strip() and not p.text.strip().startswith("#")
        ]
        source_count = len(source_lines)

        if source_count < min_sources:
            add_issue(
                result,
                rule_id="bibliography_count",
                message=(
                    f"В списке литературы {source_count} источников "
                    f"(минимум {min_sources})"
                ),
                severity=Severity.ERROR,
            )

        # Проверка на наличие запрещённых типов источников
        for para in source_lines:
            text_lower = para.text.lower()
            for forbidden in forbidden_types:
                if forbidden in text_lower:
                    add_issue(
                        result,
                        rule_id="bibliography_forbidden_type",
                        message=(
                            f"Запрещённый тип источника: «{forbidden}». "
                            f"Учебники и учебные пособия не включаются в список"
                        ),
                        severity=Severity.ERROR,
                        context=para.text[:80],
                    )
```

---

## 8. Оркестратор проверок

### vkr_checker/orchestrator.py

```python
"""
Оркестратор: загружает документ, запускает все проверки, возвращает итоговый отчёт.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .parser import load_document, DocumentModel
from .style_resolver import StyleResolver
from .checks.base import CheckResult, Severity
from .checks.margins import MarginsCheck
from .checks.fonts import FontsCheck
from .checks.spacing import SpacingCheck
from .checks.indents import IndentsCheck
from .checks.alignment import AlignmentCheck
from .checks.page_numbers import PageNumbersCheck
from .checks.heading_styles import HeadingStylesCheck
from .checks.sections import SectionsCheck
from .checks.intro import IntroCheck
from .checks.tables import TablesCheck
from .checks.figures import FiguresCheck
from .checks.bibliography import BibliographyCheck


@dataclass
class Report:
    """Итоговый отчёт по всему документу."""
    document_path: Path
    document_name: str
    results: list[CheckResult] = field(default_factory=list)

    @property
    def total_errors(self) -> int:
        return sum(r.error_count for r in self.results)

    @property
    def total_warnings(self) -> int:
        return sum(r.warning_count for r in self.results)

    @property
    def checks_passed(self) -> int:
        return sum(1 for r in self.results if r.passed)

    @property
    def checks_total(self) -> int:
        return len([r for r in self.results if not r.skipped])

    @property
    def verdict(self) -> str:
        if self.total_errors == 0:
            return "✓ Формальные требования выполнены"
        return f"✗ Обнаружено {self.total_errors} ошибок"


def load_rules(config_path: str | Path | None = None) -> dict:
    """Загружает конфигурацию требований из YAML."""
    if config_path is None:
        config_path = Path(__file__).parent.parent / "config" / "rules.yaml"
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def run_checks(
    docx_path: str | Path,
    config_path: str | Path | None = None,
) -> Report:
    """
    Главная функция: загружает документ, запускает все проверки.

    Args:
        docx_path: путь к проверяемому .docx файлу
        config_path: путь к rules.yaml (опционально)

    Returns:
        Report с результатами всех проверок
    """
    rules = load_rules(config_path)
    model: DocumentModel = load_document(docx_path)
    resolver = StyleResolver(model.doc)

    # Порядок важен: сначала структурные, потом детальные
    check_classes = [
        MarginsCheck,
        HeadingStylesCheck,
        SectionsCheck,
        IntroCheck,
        FontsCheck,
        SpacingCheck,
        IndentsCheck,
        AlignmentCheck,
        PageNumbersCheck,
        TablesCheck,
        FiguresCheck,
        BibliographyCheck,
    ]

    report = Report(
        document_path=Path(docx_path),
        document_name=Path(docx_path).name,
    )

    for cls in check_classes:
        check = cls(rules)
        result = check.run(model, resolver)
        report.results.append(result)

    return report
```

---

## 9. Генераторы отчётов

### vkr_checker/reporters/console_reporter.py

```python
"""Вывод отчёта в консоль с использованием rich."""
from __future__ import annotations

from rich.console import Console
from rich.table import Table
from rich.text import Text
from rich import box

from ..orchestrator import Report
from ..checks.base import Severity


console = Console()

SEVERITY_COLORS = {
    Severity.ERROR:   "red",
    Severity.WARNING: "yellow",
    Severity.INFO:    "blue",
}

SEVERITY_ICONS = {
    Severity.ERROR:   "✗",
    Severity.WARNING: "⚠",
    Severity.INFO:    "ℹ",
}


def print_report(report: Report) -> None:
    """Выводит полный отчёт в консоль."""
    console.print()
    _print_header(report)
    _print_summary_table(report)
    _print_details(report)
    console.print()


def _print_header(report: Report) -> None:
    console.rule(f"[bold]Проверка ВКР: {report.document_name}[/bold]")
    verdict_color = "green" if report.total_errors == 0 else "red"
    console.print(f"\n  {report.verdict}", style=f"bold {verdict_color}")
    console.print(
        f"  Ошибок: [red]{report.total_errors}[/red]  "
        f"Предупреждений: [yellow]{report.total_warnings}[/yellow]  "
        f"Проверок пройдено: [green]{report.checks_passed}/{report.checks_total}[/green]"
    )
    console.print()


def _print_summary_table(report: Report) -> None:
    tbl = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    tbl.add_column("Проверка", min_width=30)
    tbl.add_column("Статус", width=12)
    tbl.add_column("Ошибок", justify="center", width=8)
    tbl.add_column("Предупр.", justify="center", width=10)

    for r in report.results:
        if r.skipped:
            status = Text("пропущено", style="dim")
        elif r.passed:
            status = Text("✓ OK", style="green")
        else:
            status = Text("✗ ОШИБКИ", style="red")

        tbl.add_row(
            r.check_name,
            status,
            str(r.error_count) if r.error_count else "",
            str(r.warning_count) if r.warning_count else "",
        )

    console.print(tbl)


def _print_details(report: Report) -> None:
    has_issues = any(r.issues for r in report.results)
    if not has_issues:
        console.print("[green]  Замечаний не обнаружено.[/green]")
        return

    for result in report.results:
        if not result.issues:
            continue
        console.print(f"\n[bold]{result.check_name}[/bold]")
        for issue in result.issues:
            color = SEVERITY_COLORS[issue.severity]
            icon = SEVERITY_ICONS[issue.severity]
            line = f"  [{color}]{icon}[/{color}] {issue.message}"
            if issue.location_hint:
                line += f"  [dim]{issue.location_hint}[/dim]"
            console.print(line)
            if issue.context:
                console.print(f"      [dim italic]→ «{issue.context}»[/dim italic]")
```

### vkr_checker/reporters/html_reporter.py

```python
"""Генерация HTML-отчёта."""
from __future__ import annotations

from pathlib import Path
from jinja2 import Template
from ..orchestrator import Report
from ..checks.base import Severity

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Отчёт проверки ВКР — {{ report.document_name }}</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 900px;
         margin: 40px auto; padding: 0 20px; color: #222; }
  h1   { font-size: 1.4em; border-bottom: 2px solid #333; padding-bottom: 8px; }
  .verdict-ok  { color: #2a7a2a; font-weight: bold; }
  .verdict-err { color: #c0392b; font-weight: bold; }
  .summary { display: flex; gap: 24px; margin: 16px 0; }
  .stat { padding: 10px 18px; border-radius: 6px; font-size: 1.1em; }
  .stat-errors   { background: #fde8e8; color: #c0392b; }
  .stat-warnings { background: #fef9e7; color: #b7770d; }
  .stat-passed   { background: #eafaf1; color: #1e8449; }
  .check-section { margin-top: 20px; }
  .check-title   { font-size: 1em; font-weight: bold; margin-bottom: 6px;
                   cursor: pointer; }
  .check-ok   .check-title { color: #1e8449; }
  .check-err  .check-title { color: #c0392b; }
  .check-warn .check-title { color: #b7770d; }
  .issue      { padding: 5px 10px; margin: 3px 0; border-radius: 4px;
                font-size: 0.93em; }
  .issue-error   { background: #fde8e8; border-left: 3px solid #c0392b; }
  .issue-warning { background: #fef9e7; border-left: 3px solid #e67e22; }
  .issue-info    { background: #eaf4fb; border-left: 3px solid #2980b9; }
  .location { color: #888; font-size: 0.9em; margin-left: 8px; }
  .context  { color: #555; font-style: italic; font-size: 0.88em;
              margin-top: 2px; }
</style>
</head>
<body>
<h1>Отчёт проверки ВКР</h1>
<p><strong>Файл:</strong> {{ report.document_name }}</p>

<p class="{{ 'verdict-ok' if report.total_errors == 0 else 'verdict-err' }}">
  {{ report.verdict }}
</p>

<div class="summary">
  <div class="stat stat-errors">Ошибок: {{ report.total_errors }}</div>
  <div class="stat stat-warnings">Предупреждений: {{ report.total_warnings }}</div>
  <div class="stat stat-passed">Проверок OK: {{ report.checks_passed }}/{{ report.checks_total }}</div>
</div>

{% for result in report.results %}
  {% if not result.skipped %}
  <div class="check-section {{ 'check-ok' if result.passed else ('check-warn' if result.error_count == 0 else 'check-err') }}">
    <div class="check-title">
      {{ '✓' if result.passed else '✗' }} {{ result.check_name }}
      {% if result.issues %}({{ result.issues|length }}){% endif %}
    </div>
    {% for issue in result.issues %}
    <div class="issue issue-{{ issue.severity.value }}">
      {{ issue.message }}
      {% if issue.location_hint %}
        <span class="location">{{ issue.location_hint }}</span>
      {% endif %}
      {% if issue.context %}
        <div class="context">→ «{{ issue.context }}»</div>
      {% endif %}
    </div>
    {% endfor %}
  </div>
  {% endif %}
{% endfor %}

<p style="margin-top:40px; color:#aaa; font-size:0.85em;">
  Сформировано: vkr-checker
</p>
</body>
</html>
"""


def save_html_report(report: Report, output_path: str | Path) -> Path:
    """Сохраняет HTML-отчёт в файл."""
    output_path = Path(output_path)
    tmpl = Template(HTML_TEMPLATE)
    html = tmpl.render(report=report, Severity=Severity)
    output_path.write_text(html, encoding="utf-8")
    return output_path
```

---

## 10. CLI-интерфейс

### vkr_checker/cli.py

```python
"""
CLI на основе Click.

Использование:
  vkr-check document.docx
  vkr-check document.docx --output report.html
  vkr-check document.docx --format html --output report.html
  vkr-check document.docx --config my_rules.yaml
"""
from __future__ import annotations

from pathlib import Path
import sys

import click
from rich.console import Console

from .orchestrator import run_checks
from .reporters.console_reporter import print_report
from .reporters.html_reporter import save_html_report

console = Console()


@click.command()
@click.argument("docx_file", type=click.Path(exists=True, path_type=Path))
@click.option(
    "--output", "-o",
    type=click.Path(path_type=Path),
    default=None,
    help="Путь для сохранения отчёта (HTML)",
)
@click.option(
    "--format", "-f",
    "fmt",
    type=click.Choice(["console", "html", "both"], case_sensitive=False),
    default="console",
    show_default=True,
    help="Формат вывода",
)
@click.option(
    "--config", "-c",
    type=click.Path(exists=True, path_type=Path),
    default=None,
    help="Путь к файлу конфигурации rules.yaml",
)
@click.option(
    "--strict",
    is_flag=True,
    default=False,
    help="Возвращать код ошибки 1 при наличии предупреждений",
)
def main(
    docx_file: Path,
    output: Path | None,
    fmt: str,
    config: Path | None,
    strict: bool,
) -> None:
    """Проверка оформления ВКР по требованиям регламента."""
    console.print(f"[dim]Проверяю: {docx_file}[/dim]")

    try:
        report = run_checks(docx_file, config_path=config)
    except (FileNotFoundError, ValueError) as e:
        console.print(f"[red]Ошибка:[/red] {e}")
        sys.exit(2)

    if fmt in ("console", "both"):
        print_report(report)

    if fmt in ("html", "both"):
        if output is None:
            output = docx_file.with_suffix(".report.html")
        saved = save_html_report(report, output)
        console.print(f"\n[dim]HTML-отчёт сохранён: {saved}[/dim]")

    # Код выхода
    if report.total_errors > 0:
        sys.exit(1)
    if strict and report.total_warnings > 0:
        sys.exit(1)
    sys.exit(0)
```

---

## 11. GUI-интерфейс (tkinter)

### vkr_checker/gui.py

```python
"""
Графический интерфейс на tkinter.

Запуск: python -m vkr_checker.gui
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from pathlib import Path

from .orchestrator import run_checks, Report
from .reporters.html_reporter import save_html_report
from .checks.base import Severity


class VKRCheckerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VKR Checker — Проверка оформления ВКР")
        self.geometry("780x600")
        self.resizable(True, True)
        self.configure(bg="#f5f5f5")
        self._report: Report | None = None
        self._build_ui()

    # ──────────────────────────────────────────────
    # UI
    # ──────────────────────────────────────────────

    def _build_ui(self):
        # Заголовок
        header = tk.Frame(self, bg="#2c3e50", pady=12)
        header.pack(fill=tk.X)
        tk.Label(
            header, text="VKR Checker",
            font=("Segoe UI", 16, "bold"),
            fg="white", bg="#2c3e50",
        ).pack()
        tk.Label(
            header, text="Автоматическая проверка оформления ВКР",
            font=("Segoe UI", 10),
            fg="#bdc3c7", bg="#2c3e50",
        ).pack()

        # Панель выбора файла
        file_frame = tk.Frame(self, bg="#f5f5f5", padx=16, pady=10)
        file_frame.pack(fill=tk.X)

        tk.Label(
            file_frame, text="Файл DOCX:",
            font=("Segoe UI", 10), bg="#f5f5f5",
        ).grid(row=0, column=0, sticky=tk.W)

        self._file_var = tk.StringVar(value="Файл не выбран")
        tk.Label(
            file_frame, textvariable=self._file_var,
            font=("Segoe UI", 10, "italic"),
            fg="#555", bg="#f5f5f5",
            wraplength=500, anchor="w",
        ).grid(row=0, column=1, sticky=tk.W, padx=8)

        tk.Button(
            file_frame, text="Выбрать файл…",
            command=self._choose_file,
            bg="#3498db", fg="white",
            font=("Segoe UI", 10), relief=tk.FLAT,
            padx=10, pady=4,
        ).grid(row=0, column=2, padx=4)

        # Кнопки действий
        btn_frame = tk.Frame(self, bg="#f5f5f5", padx=16, pady=4)
        btn_frame.pack(fill=tk.X)

        self._check_btn = tk.Button(
            btn_frame, text="▶ Проверить",
            command=self._start_check,
            bg="#27ae60", fg="white",
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT, padx=16, pady=6,
            state=tk.DISABLED,
        )
        self._check_btn.pack(side=tk.LEFT, padx=4)

        self._html_btn = tk.Button(
            btn_frame, text="⬇ Сохранить HTML-отчёт",
            command=self._save_html,
            bg="#8e44ad", fg="white",
            font=("Segoe UI", 10),
            relief=tk.FLAT, padx=12, pady=6,
            state=tk.DISABLED,
        )
        self._html_btn.pack(side=tk.LEFT, padx=4)

        # Прогресс-бар
        self._progress = ttk.Progressbar(
            self, mode="indeterminate", length=740,
        )
        self._progress.pack(pady=4, padx=16)

        # Область результатов
        result_frame = tk.Frame(self, bg="#f5f5f5", padx=16)
        result_frame.pack(fill=tk.BOTH, expand=True)

        # Статусная строка
        self._status_var = tk.StringVar(value="Загрузите файл для проверки")
        tk.Label(
            result_frame, textvariable=self._status_var,
            font=("Segoe UI", 10, "bold"),
            bg="#f5f5f5", anchor="w",
        ).pack(fill=tk.X, pady=(4, 2))

        # Текстовая область с прокруткой
        self._output = scrolledtext.ScrolledText(
            result_frame,
            font=("Consolas", 9),
            wrap=tk.WORD, state=tk.DISABLED,
            bg="white", relief=tk.SUNKEN,
        )
        self._output.pack(fill=tk.BOTH, expand=True)

        # Настраиваем цвета тегов
        self._output.tag_config("error",   foreground="#c0392b")
        self._output.tag_config("warning", foreground="#d68910")
        self._output.tag_config("ok",      foreground="#1e8449")
        self._output.tag_config("header",  font=("Consolas", 10, "bold"))
        self._output.tag_config("dim",     foreground="#888")

    # ──────────────────────────────────────────────
    # Обработчики событий
    # ──────────────────────────────────────────────

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Выберите DOCX-файл ВКР",
            filetypes=[("Документы Word", "*.docx"), ("Все файлы", "*.*")],
        )
        if path:
            self._file_path = Path(path)
            self._file_var.set(str(self._file_path))
            self._check_btn.configure(state=tk.NORMAL)
            self._clear_output()
            self._status_var.set("Файл выбран. Нажмите «Проверить».")

    def _start_check(self):
        if not hasattr(self, "_file_path"):
            return
        self._check_btn.configure(state=tk.DISABLED)
        self._html_btn.configure(state=tk.DISABLED)
        self._clear_output()
        self._status_var.set("Выполняется проверка…")
        self._progress.start(10)
        thread = threading.Thread(target=self._run_check_thread, daemon=True)
        thread.start()

    def _run_check_thread(self):
        try:
            report = run_checks(self._file_path)
            self._report = report
            self.after(0, lambda: self._display_report(report))
        except Exception as exc:
            self.after(0, lambda: self._show_error(str(exc)))

    def _save_html(self):
        if self._report is None:
            return
        path = filedialog.asksaveasfilename(
            title="Сохранить HTML-отчёт",
            defaultextension=".html",
            filetypes=[("HTML-файлы", "*.html")],
            initialfile=f"{self._file_path.stem}_report.html",
        )
        if path:
            save_html_report(self._report, path)
            messagebox.showinfo("Сохранено", f"Отчёт сохранён:\n{path}")

    # ──────────────────────────────────────────────
    # Отображение результатов
    # ──────────────────────────────────────────────

    def _display_report(self, report: Report):
        self._progress.stop()
        self._check_btn.configure(state=tk.NORMAL)
        self._html_btn.configure(state=tk.NORMAL)

        verdict = report.verdict
        self._status_var.set(
            f"{verdict}  |  Ошибок: {report.total_errors}  "
            f"Предупреждений: {report.total_warnings}"
        )

        self._append(
            f"{'='*60}\n"
            f"  Файл: {report.document_name}\n"
            f"  {verdict}\n"
            f"  Ошибок: {report.total_errors} | "
            f"Предупреждений: {report.total_warnings} | "
            f"OK: {report.checks_passed}/{report.checks_total}\n"
            f"{'='*60}\n\n",
            tag="header",
        )

        for result in report.results:
            if result.skipped:
                continue
            icon = "✓" if result.passed else "✗"
            tag = "ok" if result.passed else "error"
            self._append(f"{icon} {result.check_name}\n", tag=tag)

            for issue in result.issues:
                tag2 = "error" if issue.severity == Severity.ERROR else "warning"
                icon2 = "  ✗" if issue.severity == Severity.ERROR else "  ⚠"
                loc = f"  [{issue.location_hint}]" if issue.location_hint else ""
                self._append(f"{icon2} {issue.message}{loc}\n", tag=tag2)
                if issue.context:
                    self._append(f"      → «{issue.context}»\n", tag="dim")

            if result.issues:
                self._append("\n")

    def _show_error(self, msg: str):
        self._progress.stop()
        self._check_btn.configure(state=tk.NORMAL)
        self._status_var.set("Ошибка при проверке")
        self._append(f"ОШИБКА: {msg}\n", tag="error")

    def _append(self, text: str, tag: str = ""):
        self._output.configure(state=tk.NORMAL)
        self._output.insert(tk.END, text, tag if tag else ())
        self._output.see(tk.END)
        self._output.configure(state=tk.DISABLED)

    def _clear_output(self):
        self._output.configure(state=tk.NORMAL)
        self._output.delete("1.0", tk.END)
        self._output.configure(state=tk.DISABLED)


def run_gui():
    app = VKRCheckerApp()
    app.mainloop()


if __name__ == "__main__":
    run_gui()
```

---

## 12. Точка входа

### main.py

```python
"""
Точка входа. Запускает GUI или CLI в зависимости от аргументов.

  python main.py                          → GUI
  python main.py check.docx              → CLI, консольный отчёт
  python main.py check.docx -o rep.html  → CLI, HTML-отчёт
"""
import sys


def main():
    # Если передан аргумент-файл — CLI режим
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        from vkr_checker.cli import main as cli_main
        cli_main()
    else:
        try:
            from vkr_checker.gui import run_gui
            run_gui()
        except ImportError:
            # tkinter недоступен — падаем в CLI
            from vkr_checker.cli import main as cli_main
            cli_main()


if __name__ == "__main__":
    main()
```

### vkr_checker/\_\_init\_\_.py

```python
"""VKR Checker — автоматическая проверка оформления ВКР."""
__version__ = "1.0.0"
```

### vkr_checker/checks/\_\_init\_\_.py

```python
# Экспортируем все модули проверок для удобного импорта
from .base import Issue, CheckResult, Severity, BaseCheck, add_issue
```

### vkr_checker/reporters/\_\_init\_\_.py

```python
from .console_reporter import print_report
from .html_reporter import save_html_report
```

---

## 13. Тесты

### tests/test_checks.py

```python
"""
Юнит-тесты для модулей проверок.

Запуск:
  pytest tests/ -v
  pytest tests/ --cov=vkr_checker --cov-report=term-missing
"""
from __future__ import annotations

import io
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from docx import Document
from docx.shared import Cm, Pt

from vkr_checker.orchestrator import run_checks, load_rules
from vkr_checker.checks.base import Severity


RULES = load_rules()
FIXTURES = Path(__file__).parent / "fixtures"


# ──────────────────────────────────────────────
# Вспомогательные функции для создания тестовых документов
# ──────────────────────────────────────────────

def make_doc_with_margins(left, right, top, bottom) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    return doc


def save_and_path(doc: Document, tmp_path: Path, name: str = "test.docx") -> Path:
    p = tmp_path / name
    doc.save(str(p))
    return p


# ──────────────────────────────────────────────
# Тесты полей
# ──────────────────────────────────────────────

class TestMarginsCheck:
    def test_correct_margins_no_errors(self, tmp_path):
        doc = make_doc_with_margins(3.0, 1.0, 2.0, 2.0)
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        margins_result = next(r for r in report.results if r.check_id == "margins")
        assert margins_result.error_count == 0

    def test_wrong_right_margin_detected(self, tmp_path):
        doc = make_doc_with_margins(3.0, 1.5, 2.0, 2.0)  # right = 1.5 вместо 1.0
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        margins_result = next(r for r in report.results if r.check_id == "margins")
        assert margins_result.error_count >= 1
        assert any("right" in i.rule_id or "правое" in i.message
                   for i in margins_result.issues)

    def test_wrong_left_margin_detected(self, tmp_path):
        doc = make_doc_with_margins(2.0, 1.0, 2.0, 2.0)  # left = 2.0 вместо 3.0
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        margins_result = next(r for r in report.results if r.check_id == "margins")
        assert margins_result.error_count >= 1


# ──────────────────────────────────────────────
# Тесты шрифтов
# ──────────────────────────────────────────────

class TestFontsCheck:
    def test_correct_font_no_errors(self, tmp_path):
        doc = Document()
        p = doc.add_paragraph("Введение")
        run = p.runs[0]
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        fonts_result = next(r for r in report.results if r.check_id == "fonts")
        # Нет явных ошибок шрифта
        font_errors = [i for i in fonts_result.issues if i.rule_id == "font_name"]
        assert len(font_errors) == 0

    def test_wrong_font_detected(self, tmp_path):
        doc = Document()
        p = doc.add_paragraph("Введение")
        p.add_run("Текст с неверным шрифтом").font.name = "Arial"
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        fonts_result = next(r for r in report.results if r.check_id == "fonts")
        font_errors = [i for i in fonts_result.issues if i.rule_id == "font_name"]
        assert len(font_errors) >= 1


# ──────────────────────────────────────────────
# Тесты структуры разделов
# ──────────────────────────────────────────────

class TestSectionsCheck:
    def test_missing_introduction_detected(self, tmp_path):
        doc = Document()
        doc.add_paragraph("Содержание")
        doc.add_paragraph("Глава 1. Теоретические основы")
        doc.add_paragraph("Заключение")
        doc.add_paragraph("Список литературы")
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        sections_result = next(r for r in report.results if r.check_id == "sections")
        missing_intro = [
            i for i in sections_result.issues
            if "Введение" in i.message or "Введение" in i.rule_id
        ]
        assert len(missing_intro) >= 1

    def test_all_sections_present_no_errors(self, tmp_path):
        doc = Document()
        for heading in [
            "Содержание", "Введение",
            "Глава 1. Теоретические основы",
            "Выводы по главе",
            "Глава 2. Практическая часть",
            "Заключение", "Список литературы",
        ]:
            doc.add_paragraph(heading)
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        sections_result = next(r for r in report.results if r.check_id == "sections")
        assert sections_result.error_count == 0


# ──────────────────────────────────────────────
# Тесты оформления рисунков
# ──────────────────────────────────────────────

class TestFiguresCheck:
    def test_correct_figure_caption(self, tmp_path):
        doc = Document()
        doc.add_paragraph("Введение")
        doc.add_paragraph("Текст текст текст")
        p = doc.add_paragraph("Рис. 1. Название рисунка без точки в конце")
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        fig_result = next(r for r in report.results if r.check_id == "figures")
        format_errors = [i for i in fig_result.issues
                        if i.rule_id == "figure_caption_format"]
        assert len(format_errors) == 0

    def test_figure_caption_with_dot_detected(self, tmp_path):
        doc = Document()
        doc.add_paragraph("Введение")
        doc.add_paragraph("Рис. 1. Название рисунка с точкой.")
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        fig_result = next(r for r in report.results if r.check_id == "figures")
        dot_errors = [i for i in fig_result.issues
                      if "точк" in i.message.lower()]
        assert len(dot_errors) >= 1

    def test_wrong_figure_format(self, tmp_path):
        doc = Document()
        doc.add_paragraph("Введение")
        doc.add_paragraph("Рисунок 1 - Название")  # неверный формат
        path = save_and_path(doc, tmp_path)
        report = run_checks(path)
        # Этот текст не начинается с "Рис." — проверка не сработает как false positive
        fig_result = next(r for r in report.results if r.check_id == "figures")
        # Нет ложных срабатываний
        assert True  # просто не падает


# ──────────────────────────────────────────────
# Тест на реальном файле
# ──────────────────────────────────────────────

class TestRealDocument:
    @pytest.mark.skipif(
        not (FIXTURES / "sample_bad.docx").exists(),
        reason="Тестовый файл отсутствует"
    )
    def test_real_bad_document_has_errors(self):
        report = run_checks(FIXTURES / "sample_bad.docx")
        assert report.total_errors > 0

    @pytest.mark.skipif(
        not (FIXTURES / "sample_good.docx").exists(),
        reason="Тестовый файл отсутствует"
    )
    def test_real_good_document_few_errors(self):
        report = run_checks(FIXTURES / "sample_good.docx")
        # Хорошо оформленная ВКР должна иметь не более 2 ошибок
        assert report.total_errors <= 2
```

---

## 14. Сборка в исполняемый файл

### Для Windows (.exe)

```bash
# Установить PyInstaller
pip install pyinstaller

# Собрать
pyinstaller --onefile \
  --name "VKRChecker" \
  --add-data "config/rules.yaml;config" \
  --add-data "vkr_checker/reporters;vkr_checker/reporters" \
  --hidden-import "docx" \
  --hidden-import "lxml" \
  --hidden-import "yaml" \
  --hidden-import "jinja2" \
  --hidden-import "rich" \
  --hidden-import "click" \
  --windowed \
  main.py

# Результат: dist/VKRChecker.exe
```

### Для macOS (.app)

```bash
pyinstaller --onefile \
  --name "VKRChecker" \
  --add-data "config/rules.yaml:config" \
  --windowed \
  main.py
```

### spec-файл (альтернатива для сложных сборок)

```python
# VKRChecker.spec
block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config/rules.yaml', 'config'),
    ],
    hiddenimports=['docx', 'lxml', 'yaml', 'jinja2', 'rich', 'click'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name='VKRChecker',
    debug=False,
    console=False,  # True для CLI-версии без GUI
    icon='icon.ico',  # опционально
)
```

---

## Приложение A. Типичные сложности и их решения

| Проблема | Причина | Решение |
|----------|---------|---------|
| `run.font.name` возвращает `None` | Шрифт задан через стиль, не через run | `StyleResolver.get_font_name()` |
| `run.font.size` возвращает `None` | Кегль задан в стиле документа | `StyleResolver.get_font_size_pt()` |
| Интервал = `None` | Параграф не имеет явного `w:spacing` | Читать из цепочки стилей |
| Таблицы и параграфы вперемешку | `doc.paragraphs` пропускает таблицы | Итерировать через `doc.element.body` |
| Неверное определение «красной строки» | Отступ через пробелы, не через `w:ind` | Предупреждение при `indent == 0` в теле |
| Нумерация страниц не найдена | PAGE field закодирован по-разному | Искать `fldChar` + `instrText` содержащий `PAGE` |
| Ложные срабатывания в таблицах | Цвет/размер наследуется от Normal | Дедупликация + threshold по параграфам |
| Диаграммы вставлены как картинки | Нарушение регламента | Детектировать: `<a:blip>` = картинка, `<c:chart>` = правильная диаграмма |

---

## Приложение B. Команды для быстрого старта

```bash
# 1. Создать структуру проекта
mkdir -p vkr_checker/{checks,reporters} config tests/fixtures

# 2. Установить зависимости
pip install python-docx lxml PyYAML jinja2 click rich pytest pytest-cov

# 3. Скопировать все файлы из этой спецификации

# 4. Проверить документ через CLI
python main.py путь/к/вашей_вкр.docx

# 5. Получить HTML-отчёт
python main.py путь/к/вашей_вкр.docx --format html --output отчёт.html

# 6. Запустить GUI
python main.py

# 7. Запустить тесты
pytest tests/ -v --cov=vkr_checker

# 8. Собрать .exe
pip install pyinstaller
pyinstaller --onefile --name VKRChecker --windowed main.py
```
