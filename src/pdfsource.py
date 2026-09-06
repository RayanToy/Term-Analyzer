# -*- coding: utf-8 -*-
"""Извлечение из PDF того текста, которого нет в .docx.

Сверка показала, что PDF-версии учебников содержат материал, отсутствующий
в .docx: подписи к иллюстрациям и вопросы к ним, титульный лист,
колонтитулы, условные обозначения. Основной подсчёт идёт по .docx (по нему
же делается подсветка), а этот модуль выделяет «добавку» PDF, чтобы по ней
можно было отдельно посчитать единицы культурного кода.

Текст из PDF заведомо грязный: сбитый порядок колонок, служебные токены
вроде «0A21B», разорванные строки. Поэтому результат идёт отдельным листом
с пометкой «требует проверки», а не в общие числа.
"""

from __future__ import annotations

import difflib
import re
from pathlib import Path

#: перенос слова со строки на строку: «эпо- хой» -> «эпохой»
HYPHEN_BREAK = re.compile(r"(\w)[-‐‑]\s+([а-яё]\w*)")

#: голый номер страницы
PAGE_NUMBER = re.compile(r"^\d+$")

#: служебные метки вида «0A21B», которые оставляет извлечение текста
ARTIFACT = re.compile(r"^[0-9A-F]{4,8}$")

#: сегменты короче этого числа слов не берём — там нечего искать
MIN_SEGMENT_WORDS = 3


def read_pdf_words(path: Path) -> list[str]:
    """Текст PDF как список слов: переносы склеены, номера страниц убраны."""
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    chunks = []
    for page in reader.pages:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:  # pragma: no cover — отдельные битые страницы
            chunks.append("")

    text = " ".join(chunks)
    previous = None
    while previous != text:
        previous = text
        text = HYPHEN_BREAK.sub(r"\1\2", text)

    return [
        w
        for w in text.split()
        if not PAGE_NUMBER.match(w) and not ARTIFACT.match(w)
    ]


#: сегмент, повторяющийся столько раз и чаще, считается колонтитулом
REPEAT_THRESHOLD = 3

#: типографские метки, которые тянутся из макета на каждой странице
#: («15-1746-05-003-496_o 12 . indd 22.06.2026 19 : 24»). Из-за меняющегося
#: времени одинаковые колонтитулы переставали схлопываться.
PRINT_MARKS = re.compile(
    r"\S*\.\s*indd\b|\b\d{2}\.\d{2}\.\d{4}\b|\b\d{1,2}\s*:\s*\d{2}\s*:\s*\d{2}\b"
)

#: титул и выходные данные: место издания, министерство, УДК/ББК/ISBN,
#: тираж, редколлегия. Это не текст учебника, а обложка и колофон, и
#: «МОСКВА • ПРОСВЕЩЕНИЕ» давало 374 ложных вхождения единицы «Москва»,
#: «Минпросвещения России» — 256 вхождений «Россия», имя автора в строке
#: «Мединский Владимир Ростиславович» — 36 вхождений «Владимир».
IMPRINT = re.compile(
    r"УДК|ББК|ISBN|Учебное издание|Минпросвещения|Исключительные права"
    r"|Ответственн\w+ (?:редактор|секретар)|Подписано в печать|Тираж"
    r"|Гарнитура|усл\.\s*печ|ПРОСВЕЩЕНИЕ\s*•|Художественн\w+ редактор"
    r"|Технический редактор|Корректор|Вёрстка|Верстка"
)


def pdf_only_segments(pdf_path: Path, docx_words: list[str]) -> list[str]:
    """Куски PDF, которых нет в .docx.

    Сравнение пословное; каждый несовпавший блок становится отдельным
    «абзацем» — так у найденных единиц будет вменяемый контекст.

    Колонтитулы («ИСТОРИЯ РОССИИ» вверху каждой страницы) попадают сюда по
    одному разу на страницу и накручивают сотни ложных вхождений, поэтому
    повторяющиеся сегменты оставляем только один раз.
    """
    pdf = read_pdf_words(pdf_path)
    matcher = difflib.SequenceMatcher(None, docx_words, pdf, autojunk=False)

    raw = []
    for tag, _i1, _i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        words = pdf[j1:j2]
        if len(words) < MIN_SEGMENT_WORDS:
            continue
        segment = " ".join(words)
        if IMPRINT.search(segment):
            # титул и колофон — не текст учебника
            continue
        segment = PRINT_MARKS.sub(" ", segment)
        segment = " ".join(segment.split())
        if len(segment.split()) >= MIN_SEGMENT_WORDS:
            raw.append(segment)

    seen: dict[str, int] = {}
    for segment in raw:
        key = segment.casefold()
        seen[key] = seen.get(key, 0) + 1

    kept: set[str] = set()
    segments = []
    for segment in raw:
        key = segment.casefold()
        if seen[key] >= REPEAT_THRESHOLD:
            if key in kept:
                continue
            kept.add(key)
        segments.append(segment)
    return segments


GRADE_IN_PDF = re.compile(r"-(\d{1,2})-klass")
GRADE_IN_DOCX = re.compile(r"_Hs_(\d{2})_")


def grade_of(path: Path) -> str | None:
    """Номер класса из имени файла (форматы .pdf и .docx различаются)."""
    for pattern in (GRADE_IN_PDF, GRADE_IN_DOCX):
        m = pattern.search(path.name)
        if m:
            return str(int(m.group(1)))
    return None


def pair_pdfs_with_docx(books: list[Path]) -> list[tuple[Path, Path]]:
    """Сопоставляет PDF и .docx по разделу и классу."""
    by_key = {(b.parent.name, grade_of(b)): b for b in books}
    pairs = []
    for docx in books:
        section = docx.parent
        grade = grade_of(docx)
        for pdf in sorted(section.glob("*.pdf")):
            if grade_of(pdf) == grade and by_key.get((section.name, grade)) == docx:
                pairs.append((pdf, docx))
    return pairs
