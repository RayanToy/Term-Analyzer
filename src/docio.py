# -*- coding: utf-8 -*-
"""Чтение docx: абзацы в порядке документа вместе с их run'ами.

Текст абзаца собирается как конкатенация текстов всех ``w:r`` (включая те,
что лежат внутри гиперссылок), поэтому символьные позиции токенов всегда
совпадают с позициями внутри run'ов — это то, на чём держится подсветка.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from docx import Document
from docx.oxml.ns import qn
from docx.text.run import Run


@dataclass
class ParagraphView:
    """Абзац вместе с плоским списком его run'ов и общим текстом."""

    paragraph: object
    runs: list = field(default_factory=list)
    text: str = ""

    @classmethod
    def build(cls, paragraph) -> "ParagraphView":
        run_elements = paragraph._p.xpath(".//w:r")
        runs = [Run(el, paragraph) for el in run_elements]
        return cls(paragraph=paragraph, runs=runs, text="".join(r.text for r in runs))


def _iter_block_paragraphs(parent_element, parent_obj):
    """Обходит w:p и w:tbl в порядке документа."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in parent_element.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent_obj)
        elif child.tag == qn("w:tbl"):
            table = Table(child, parent_obj)
            for row in table.rows:
                for cell in row.cells:
                    for para in _iter_block_paragraphs(cell._tc, cell):
                        yield para


def open_document(path):
    """Открывает docx."""
    return Document(str(path))


def paragraph_views(doc) -> list[ParagraphView]:
    """Все абзацы документа (включая ячейки таблиц) в порядке следования."""
    return [
        ParagraphView.build(p)
        for p in _iter_block_paragraphs(doc.element.body, doc)
    ]
