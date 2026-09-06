# -*- coding: utf-8 -*-
"""Подсветка найденных единиц в docx с сохранением форматирования.

В исходном ``apply_highlights_to_paragraph`` из Term-Analyzer абзац с
совпадением полностью пересобирался: все run'ы удалялись и создавались
новые без ``rPr``. Для учебника это неприемлемо — терялись жирный шрифт,
курсив, кегль, заголовочные стили.

Здесь вместо этого run режется по границам совпадения: каждый кусок —
глубокая копия исходного ``w:r`` со всем его ``w:rPr``, и цвет ставится
только куску, попавшему внутрь совпадения. Абзацы без совпадений не
трогаются вообще.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.text.run import Run

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

#: элементы, дающие содержимое run'а
CONTENT_TAGS = {
    qn("w:t"),
    qn("w:tab"),
    qn("w:br"),
    qn("w:cr"),
    qn("w:drawing"),
    qn("w:object"),
    qn("w:pict"),
    qn("w:noBreakHyphen"),
    qn("w:softHyphen"),
    qn("w:sym"),
    qn("w:fldChar"),
    qn("w:instrText"),
}


@dataclass
class HighlightStats:
    """Статистика применения подсветки (для контроля качества)."""

    paragraphs_touched: int = 0
    runs_split: int = 0
    spans_applied: int = 0
    spans_skipped_complex_run: int = 0


#: содержимое, которое нельзя разрезать: картинки, поля, объекты
UNSAFE_TAGS = {
    qn("w:drawing"),
    qn("w:object"),
    qn("w:pict"),
    qn("w:fldChar"),
    qn("w:instrText"),
}

#: сколько символов даёт элемент в Run.text
ATOMIC_LENGTH = {qn("w:tab"): 1, qn("w:br"): 1, qn("w:cr"): 1}


def _is_splittable(run_element) -> bool:
    """Можно ли резать run.

    Табы и переносы строк резать можно — они просто уезжают в тот кусок,
    куда попали по смещению. Нельзя трогать run'ы с картинками, полями и
    внедрёнными объектами.
    """
    return not any(c.tag in UNSAFE_TAGS for c in run_element)


def _piece_index(pieces, offset: int) -> int:
    """В какой кусок попадает символьное смещение внутри run'а."""
    for i, (start, end, _color) in enumerate(pieces):
        if start <= offset < end:
            return i
    return len(pieces) - 1


def _set_color(run_element, paragraph, hex_color: str) -> None:
    """Ставит цвет шрифта, сохраняя остальное форматирование."""
    run = Run(run_element, paragraph)
    run.font.color.rgb = RGBColor.from_string(hex_color)
    color_el = run_element.find(qn("w:rPr"))
    if color_el is not None:
        for c in color_el.findall(qn("w:color")):
            # themeColor перебивает явный w:val — убираем его
            for attr in (qn("w:themeColor"), qn("w:themeTint"), qn("w:themeShade")):
                if attr in c.attrib:
                    del c.attrib[attr]


def _split_run(run_element, paragraph, pieces, stats: HighlightStats) -> None:
    """Заменяет run набором копий: по одной на кусок ``(начало, конец, цвет)``.

    Форматирование сохраняется: каждый кусок — копия исходного run'а вместе
    с его ``w:rPr``, меняется только содержимое и цвет. Содержимое
    распределяется по кускам в порядке следования: текст режется по
    смещениям, табы и переносы целиком уходят в свой кусок.
    """
    parent = run_element.getparent()
    index = parent.index(run_element)

    # Шаблон — тот же run без содержимого (rPr и прочее остаются на месте).
    template = copy.deepcopy(run_element)
    for child in list(template):
        if child.tag in CONTENT_TAGS:
            template.remove(child)

    buckets: list[list] = [[] for _ in pieces]
    offset = 0
    for child in run_element:
        if child.tag not in CONTENT_TAGS:
            continue
        if child.tag == qn("w:t"):
            text = child.text or ""
            for i, (start, end, _color) in enumerate(pieces):
                a = max(start, offset)
                b = min(end, offset + len(text))
                if a < b:
                    part = copy.deepcopy(child)
                    part.text = text[a - offset : b - offset]
                    part.set(XML_SPACE, "preserve")
                    buckets[i].append(part)
            offset += len(text)
        else:
            buckets[_piece_index(pieces, offset)].append(copy.deepcopy(child))
            offset += ATOMIC_LENGTH.get(child.tag, 0)

    new_elements = []
    for (_start, _end, color), children in zip(pieces, buckets):
        if not children:
            continue
        clone = copy.deepcopy(template)
        for child in children:
            clone.append(child)
        if color:
            _set_color(clone, paragraph, color)
        new_elements.append(clone)

    if not new_elements:
        return

    for position, element in enumerate(new_elements):
        parent.insert(index + position, element)
    parent.remove(run_element)
    stats.runs_split += 1


def apply_spans(view, spans: list[tuple[int, int, str]], stats: HighlightStats) -> None:
    """Красит символьные интервалы абзаца.

    Args:
        view: ``ParagraphView`` из ``src.docio``
        spans: список ``(начало, конец, HEX-цвет)`` в координатах ``view.text``;
            интервалы не должны пересекаться
        stats: счётчики для отчёта
    """
    if not spans:
        return

    spans = sorted(spans)
    touched = False

    position = 0
    for run in view.runs:
        length = len(run.text)
        run_start, run_end = position, position + length
        position = run_end
        if length == 0:
            continue

        overlapping = [
            (max(s, run_start), min(e, run_end), c)
            for s, e, c in spans
            if s < run_end and e > run_start
        ]
        if not overlapping:
            continue

        if not _is_splittable(run._element):
            stats.spans_skipped_complex_run += len(overlapping)
            continue

        # Куски в координатах внутри run'а, подряд и без разрывов.
        pieces: list[tuple[int, int, str | None]] = []
        cursor = 0
        for s, e, color in overlapping:
            rel_s, rel_e = s - run_start, e - run_start
            if cursor < rel_s:
                pieces.append((cursor, rel_s, None))
            pieces.append((rel_s, rel_e, color))
            cursor = rel_e
        if cursor < length:
            pieces.append((cursor, length, None))

        if len(pieces) == 1 and pieces[0][2] is None:
            continue

        _split_run(run._element, view.paragraph, pieces, stats)
        stats.spans_applied += len(overlapping)
        touched = True

    if touched:
        stats.paragraphs_touched += 1


def resolve_overlaps(spans: list[tuple[int, int, str]], priority: dict[str, int]) -> list[tuple[int, int, str]]:
    """Разрешает пересечения интервалов подсветки.

    Пересекающиеся вхождения (например «Аврора» внутри «Крейсер "Аврора"»)
    объединяются в один интервал; цвет берётся по приоритету (уверенное
    вхождение важнее спорного).
    """
    if not spans:
        return []
    spans = sorted(spans)
    merged: list[list] = [list(spans[0])]
    for start, end, color in spans[1:]:
        last = merged[-1]
        if start < last[1]:
            last[1] = max(last[1], end)
            if priority.get(color, 0) > priority.get(last[2], 0):
                last[2] = color
        else:
            merged.append([start, end, color])
    return [(a, b, c) for a, b, c in merged]
