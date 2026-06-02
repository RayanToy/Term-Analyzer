"""
Подсветка терминов в docx-файлах.

Логика работы:
    1. Для каждого абзаца и ячейки таблицы вычисляются
       символьные интервалы [start, end) найденных терминов.
    2. Интервалы объединяются (merge_spans) во избежание наложений.
    3. Текст абзаца перезаписывается новыми runs:
       - обычный run  — текст вне терминов
       - красный run  — текст термина (RGBColor(255, 0, 0))

Важно:
    Функция работает с символьными позициями исходного текста абзаца,
    которые возвращает tokens_with_punct() через token.start / token.stop.
    Поэтому текст подаётся в tokens_with_punct() без strip() и без
    удаления пунктуации.
"""

import os
from pathlib import Path

from docx import Document
from docx.shared import RGBColor

from src.analyzer import (
    tokens_with_punct,
    build_indexed_words,
    _words_are_adjacent,
    merge_spans,          # добавим в analyzer.py (см. ниже)
)
from src.preprocessing import TextPreprocessor


# =============================================================================
# Вычисление интервалов подсветки
# =============================================================================

def compute_highlight_spans(
    text: str,
    patterns_by_len: dict[int, list[str]],
    preprocessor: TextPreprocessor,
    allow_punct_between: bool = False,
    disallow_hyphen_adjacent_for_unigrams: bool = True,
) -> list[tuple[int, int]]:
    """
    Вычисляет символьные интервалы [start, end) вхождений терминов в тексте.

    Возвращает объединённый (без пересечений) отсортированный список интервалов.
    Если текст пустой или паттернов нет — возвращает пустой список.

    Args:
        text: текст абзаца (без предобработки — нужны оригинальные позиции)
        patterns_by_len: словарь {длина_паттерна: [паттерн, ...]}
        preprocessor: экземпляр TextPreprocessor
        allow_punct_between: допускать пунктуацию между словами фразы
        disallow_hyphen_adjacent_for_unigrams: не подсвечивать унигаммы рядом с дефисом

    Returns:
        Список неперекрывающихся интервалов [(start, end), ...]
    """
    if not text or not text.strip():
        return []

    tokens = tokens_with_punct(text, preprocessor)
    words = build_indexed_words(tokens)
    spans = []

    # -----------------------------------------------------------------
    # Унигаммы
    # -----------------------------------------------------------------
    unis = set(patterns_by_len.get(1, []))
    if unis:
        for w in words:
            if w['lemma'] not in unis:
                continue
            if disallow_hyphen_adjacent_for_unigrams and (
                w['hyph_left'] or w['hyph_right'] or w['inner_hyphen']
            ):
                continue
            spans.append((w['start'], w['end']))

    # -----------------------------------------------------------------
    # k-граммы (k >= 2)
    # -----------------------------------------------------------------
    for k, pats in patterns_by_len.items():
        if k <= 1 or not pats or len(words) < k:
            continue

        patset = set(pats)

        for i in range(len(words) - k + 1):
            if not _words_are_adjacent(words, tokens, i, k, allow_punct_between):
                continue

            phrase = " ".join(w['lemma'] for w in words[i:i + k])
            if phrase not in patset:
                continue

            spans.append((words[i]['start'], words[i + k - 1]['end']))

    return merge_spans(spans)


# =============================================================================
# Применение подсветки к абзацу
# =============================================================================

def apply_highlights_to_paragraph(
    paragraph,
    spans: list[tuple[int, int]],
) -> None:
    """
    Перезаписывает runs абзаца, выделяя красным цветом указанные интервалы.

    Если spans пустой — абзац не изменяется.

    Важно: функция удаляет все существующие runs абзаца и создаёт новые.
    Форматирование (жирный, курсив и т.д.) при этом не сохраняется —
    сохраняется только цвет для найденных терминов.

    Args:
        paragraph: объект Paragraph из python-docx
        spans: список интервалов [(start, end), ...] для выделения
    """
    if not spans:
        return

    text = paragraph.text

    # Удаляем существующие runs
    for run in paragraph.runs:
        paragraph._element.remove(run._element)

    # Создаём новые runs с подсветкой
    pos = 0
    for start, end in spans:
        # Обычный текст до термина
        if pos < start:
            paragraph.add_run(text[pos:start])

        # Термин — красным
        run = paragraph.add_run(text[start:end])
        run.font.color.rgb = RGBColor(255, 0, 0)

        pos = end

    # Остаток текста после последнего термина
    if pos < len(text):
        paragraph.add_run(text[pos:])


# =============================================================================
# Применение подсветки к ячейке таблицы
# =============================================================================

def apply_highlights_to_cell(
    cell,
    patterns_by_len: dict[int, list[str]],
    preprocessor: TextPreprocessor,
    allow_punct_between: bool = False,
    disallow_hyphen_adjacent_for_unigrams: bool = True,
) -> None:
    """
    Применяет подсветку ко всем абзацам внутри ячейки таблицы.

    Args:
        cell: объект Cell из python-docx
        patterns_by_len: словарь {длина_паттерна: [паттерн, ...]}
        preprocessor: экземпляр TextPreprocessor
        allow_punct_between: допускать пунктуацию между словами фразы
        disallow_hyphen_adjacent_for_unigrams: не подсвечивать унигаммы рядом с дефисом
    """
    for paragraph in cell.paragraphs:
        spans = compute_highlight_spans(
            paragraph.text,
            patterns_by_len,
            preprocessor,
            allow_punct_between=allow_punct_between,
            disallow_hyphen_adjacent_for_unigrams=disallow_hyphen_adjacent_for_unigrams,
        )
        apply_highlights_to_paragraph(paragraph, spans)


# =============================================================================
# Аннотирование docx-файла
# =============================================================================

def annotate_docx(
    in_path: str | Path,
    out_path: str | Path,
    patterns_by_len: dict[int, list[str]],
    preprocessor: TextPreprocessor,
    allow_punct_between: bool = False,
    disallow_hyphen_adjacent_for_unigrams: bool = True,
) -> None:
    """
    Создаёт копию docx-файла с подсвеченными терминами.

    Обрабатывает:
    - все абзацы документа
    - все ячейки всех таблиц документа

    Выходной файл сохраняется в out_path.
    Директория создаётся автоматически, если не существует.

    Args:
        in_path: путь к исходному docx-файлу
        out_path: путь для сохранения аннотированного файла
        patterns_by_len: словарь {длина_паттерна: [паттерн, ...]}
        preprocessor: экземпляр TextPreprocessor
        allow_punct_between: допускать пунктуацию между словами фразы
        disallow_hyphen_adjacent_for_unigrams: не подсвечивать унигаммы рядом с дефисом
    """
    doc = Document(str(in_path))

    # Абзацы
    for paragraph in doc.paragraphs:
        spans = compute_highlight_spans(
            paragraph.text,
            patterns_by_len,
            preprocessor,
            allow_punct_between=allow_punct_between,
            disallow_hyphen_adjacent_for_unigrams=disallow_hyphen_adjacent_for_unigrams,
        )
        apply_highlights_to_paragraph(paragraph, spans)

    # Таблицы
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                apply_highlights_to_cell(
                    cell,
                    patterns_by_len,
                    preprocessor,
                    allow_punct_between=allow_punct_between,
                    disallow_hyphen_adjacent_for_unigrams=disallow_hyphen_adjacent_for_unigrams,
                )

    out_path = Path(out_path)
    os.makedirs(out_path.parent, exist_ok=True)
    doc.save(str(out_path))
    print(f"[Учебник] Сохранён: {out_path}")