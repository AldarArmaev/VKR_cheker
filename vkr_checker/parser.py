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


def load_document(
    path: str | Path,
    rules: dict | None = None,
) -> DocumentModel:
    """
    Загружает DOCX и строит DocumentModel.

    Args:
        path: путь к файлу
        rules: правила для детекции разделов (опционально)

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
    blocks = list(iter_blocks(doc))
    paragraphs = [b.paragraph for b in blocks if b.kind == "paragraph"]
    tables = [b.table for b in blocks if b.kind == "table"]
    sections = _detect_sections(paragraphs, rules)
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


def iter_blocks(doc: Document) -> Iterator[BlockItem]:
    """
    Итерирует по документу в правильном порядке, включая таблицы.
    python-docx doc.paragraphs пропускает таблицы, поэтому используем XML.
    
    Обходит doc.element.body и корректно обрабатывает:
    - w:p (параграфы)
    - w:tbl (таблицы) — с вложенным циклом по строкам (w:tr) и ячейкам (w:tc)
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
            # Это таблица — обрабатываем её целиком
            table = _find_table_by_element(doc, child)
            if table is not None:
                yield BlockItem(index=idx, kind="table", table=table)
                idx += 1


def _find_paragraph_by_element(doc: Document, elem) -> Paragraph | None:
    """Найти объект Paragraph по его XML элементу."""
    for p in doc.paragraphs:
        if p._p is elem:
            return p
    return None


def _find_table_by_element(doc: Document, elem) -> Table | None:
    """Найти объект Table по его XML элементу."""
    for t in doc.tables:
        if t._tbl is elem:
            return t
    return None


def _detect_sections(
    paragraphs: list[Paragraph],
    rules: dict | None = None,
) -> list[SectionInfo]:
    """Определяет разделы документа по тексту заголовков."""
    if rules is None:
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

    # Если «Введение» не найдено, возвращаем -1, что сигнализирует об ошибке
    # Проверки будут пропущены или выдадут предупреждение
    return intro_idx, next_major_idx
