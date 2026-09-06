# -*- coding: utf-8 -*-
"""Обработка одного учебника: токенизация -> поиск -> разрешение -> подсветка.

Один проход по абзацам даёт сразу и счётчики для Excel, и символьные
интервалы для подсветки, так что числа в отчёте и в docx не могут разойтись.
"""

from __future__ import annotations

import time
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import config
from src import analyzer, disambiguation, docio, highlighter, pdfsource, review
from src.disambiguation import (
    COUNTED,
    STATUS_ACCEPTED,
    STATUS_CONFIDENT,
    STATUS_LOWERCASE,
    STATUS_REJECTED,
)

#: приоритет цветов при пересечении интервалов
COLOR_PRIORITY = {config.COLOR_CONFIDENT: 2, config.COLOR_DISPUTED: 1}

STATUS_LABELS = {
    STATUS_CONFIDENT: "уверенное",
    STATUS_ACCEPTED: "спорное, засчитано",
    STATUS_REJECTED: "спорное, отклонено",
    STATUS_LOWERCASE: "отклонено: строчная буква",
}


@dataclass
class BookResult:
    """Результат по одному учебнику."""

    rel_path: str
    section: str
    stem: str
    backend: str
    counts: dict = field(default_factory=lambda: defaultdict(int))
    accepted: dict = field(default_factory=lambda: defaultdict(int))
    rejected: dict = field(default_factory=lambda: defaultdict(int))
    lowercase: dict = field(default_factory=lambda: defaultdict(int))
    contexts: list = field(default_factory=list)
    disputed_rows: list = field(default_factory=list)
    review_rows: list = field(default_factory=list)
    manual: dict = field(default_factory=lambda: defaultdict(int))
    pdf: dict = field(default_factory=lambda: defaultdict(int))
    match_keys: dict = field(default_factory=dict)
    total_tokens: int = 0
    total_words: int = 0
    paragraphs: int = 0
    #: (раздел, класс) — связывает результат по PDF с результатом по .docx
    pair_key: tuple | None = None
    elapsed: float = 0.0
    highlight_stats: highlighter.HighlightStats | None = None

    @property
    def total_hits(self) -> int:
        """Сумма по строкам отчёта.

        Наименование, входящее в несколько тематических групп (Алёнушка —
        ИЗО / Лит. персонажи / Фольклор), даёт отдельную строку в каждой
        группе, поэтому одно место в тексте учитывается здесь несколько раз.
        """
        return sum(self.counts.values())

    @property
    def distinct_hits(self) -> int:
        """Число уникальных мест в тексте (без повторов по группам)."""
        return len(
            {
                (para, start, end)
                for (_unit, para, start, end), status in self.match_keys.items()
                if status in COUNTED
            }
        )

    @property
    def unique_units(self) -> int:
        return sum(1 for v in self.counts.values() if v)


def process_book(
    docx_path: Path,
    units,
    index,
    first_lemmas,
    backend,
    highlight_out: Path | None = None,
    lemma_sink: dict | None = None,
    form_freq: dict | None = None,
) -> BookResult:
    """Полный проход по одному учебнику.

    ``lemma_sink`` и ``form_freq`` (если переданы) накапливают соответствие
    «словоформа -> лемма» и частоты словоформ — они нужны только для листа
    «Спорные леммы» в отчёте сравнения бэкендов.
    """
    started = time.perf_counter()

    section = docx_path.parent.name
    result = BookResult(
        rel_path=f"{section}\\{docx_path.stem}",
        section=section,
        stem=docx_path.stem,
        backend=backend.name,
        pair_key=(section, pdfsource.grade_of(docx_path)),
    )

    doc = docio.open_document(docx_path)
    views = docio.paragraph_views(doc)
    result.paragraphs = len(views)

    paragraph_tokens, all_matches = _analyze(
        [v.text for v in views], units, index, first_lemmas, backend,
        result, lemma_sink, form_freq,
    )
    _collect(result, all_matches, units)
    result.review_rows = review.collect_possible_misses(
        units,
        result.counts,
        paragraph_tokens,
        filtered=set(result.rejected) | set(result.lowercase),
    )
    _apply_manual_additions(result)

    if highlight_out is not None:
        result.highlight_stats = _highlight(
            views, all_matches, highlight_out, doc
        )

    result.elapsed = time.perf_counter() - started
    return result


def _apply_manual_additions(result: BookResult) -> None:
    """Добавляет в подсчёт пропуски, подтверждённые ручной проверкой.

    Это места, где объект действительно назван, но записан не так, как в
    списке: вставлено слово, другой порядок, перечисление с общим словом.
    Автоматически такое не отличить от случайных совпадений, поэтому решение
    принято вручную (data/review_verdicts.txt).

    Такие вхождения попадают в «Количество» и «Итого», но помечены отдельной
    колонкой и НЕ подсвечены в docx — дословно фразы там нет.
    """
    for row in result.review_rows:
        if row.get("verdict") not in config.REVIEW_COUNTED_VERDICTS:
            continue
        idx = row["unit_idx"]
        result.counts[idx] += 1
        result.manual[idx] += 1
        result.contexts.append(
            {
                "unit_idx": idx,
                "n": result.counts[idx],
                "variant": row["variant"],
                "kind": "ручная проверка",
                "found": "—",
                "context": row["context"],
                "status": "добавлено по ручной проверке (в docx не подсвечено)",
                "reason": row.get("comment", ""),
            }
        )


def _analyze(
    texts,
    units,
    index,
    first_lemmas,
    backend,
    result: BookResult,
    lemma_sink: dict | None = None,
    form_freq: dict | None = None,
):
    """Токенизирует блоки текста, ищет совпадения и разрешает спорные."""
    paragraph_tokens: list[list] = []
    all_matches: list[analyzer.Match] = []

    for para_idx, text in enumerate(texts):
        if not text.strip():
            paragraph_tokens.append([])
            continue
        tokens = backend.tokenize(text)
        paragraph_tokens.append(tokens)
        result.total_tokens += len(tokens)
        result.total_words += sum(1 for t in tokens if t.is_word)
        if lemma_sink is not None:
            for t in tokens:
                if t.is_word:
                    lemma_sink.setdefault(t.text, t.lemma)
                    if form_freq is not None:
                        form_freq[t.text] = form_freq.get(t.text, 0) + 1
        all_matches.extend(
            analyzer.find_matches(tokens, index, first_lemmas, para_idx)
        )

    disambiguation.resolve(all_matches, paragraph_tokens, units)
    return paragraph_tokens, all_matches


def _collect(result: BookResult, matches, units) -> None:
    """Раскладывает совпадения по счётчикам, контекстам и спорным."""
    per_unit_contexts: dict[int, int] = defaultdict(int)

    for m in matches:
        unit = units[m.unit_idx]
        variant = unit.variants[m.variant_idx]
        result.match_keys[(m.unit_idx, m.para_idx, m.char_start, m.char_end)] = m.status

        if m.status in COUNTED:
            result.counts[m.unit_idx] += 1
            if m.status == STATUS_ACCEPTED:
                result.accepted[m.unit_idx] += 1
            per_unit_contexts[m.unit_idx] += 1
            if per_unit_contexts[m.unit_idx] <= config.MAX_CONTEXTS_PER_UNIT:
                result.contexts.append(
                    {
                        "unit_idx": m.unit_idx,
                        "n": per_unit_contexts[m.unit_idx],
                        "variant": variant.surface,
                        "kind": variant.kind,
                        "found": m.surface,
                        "context": m.context,
                        "status": STATUS_LABELS[m.status],
                        "reason": m.reason,
                    }
                )
        elif m.status == STATUS_REJECTED:
            result.rejected[m.unit_idx] += 1
        elif m.status == STATUS_LOWERCASE:
            result.lowercase[m.unit_idx] += 1

        if m.status in (STATUS_ACCEPTED, STATUS_REJECTED):
            result.disputed_rows.append(
                {
                    "unit_idx": m.unit_idx,
                    "variant": variant.surface,
                    "found": m.surface,
                    "context": m.context,
                    "status": STATUS_LABELS[m.status],
                    "reason": m.reason,
                }
            )


def process_pdf_extra(
    pdf_path: Path, docx_path: Path, units, index, first_lemmas, backend
) -> BookResult:
    """Ищет единицы в той части PDF, которой нет в .docx.

    Это подписи к иллюстрациям, вопросы к ним, титул и колонтитулы. Числа
    идут отдельным листом: текст из PDF извлекается с искажениями, и в
    общий подсчёт его мешать нельзя.
    """
    started = time.perf_counter()
    section = pdf_path.parent.name
    result = BookResult(
        rel_path=f"{section}\\{pdf_path.stem}",
        section=section,
        stem=pdf_path.stem,
        backend=backend.name,
        pair_key=(section, pdfsource.grade_of(pdf_path)),
    )

    doc = docio.open_document(docx_path)
    docx_words = " ".join(v.text for v in docio.paragraph_views(doc)).split()
    segments = pdfsource.pdf_only_segments(pdf_path, docx_words)
    result.paragraphs = len(segments)

    _, matches = _analyze(segments, units, index, first_lemmas, backend, result)
    _collect(result, matches, units)

    result.elapsed = time.perf_counter() - started
    return result


def _highlight(views, matches, out_path: Path, doc) -> highlighter.HighlightStats:
    """Красит совпадения и сохраняет копию документа."""
    stats = highlighter.HighlightStats()

    spans_by_para: dict[int, list[tuple[int, int, str]]] = defaultdict(list)
    for m in matches:
        if m.status in COUNTED:
            color = config.COLOR_CONFIDENT
        elif m.status == STATUS_REJECTED and config.HIGHLIGHT_REJECTED:
            color = config.COLOR_DISPUTED
        else:
            continue
        spans_by_para[m.para_idx].append((m.char_start, m.char_end, color))

    for para_idx, spans in spans_by_para.items():
        merged = highlighter.resolve_overlaps(spans, COLOR_PRIORITY)
        highlighter.apply_spans(views[para_idx], merged, stats)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return stats


def merge_pdf_extra(book: BookResult, pdf_result: BookResult, units) -> None:
    """Прибавляет к учебнику вхождения из PDF-добавки.

    PDF-добавка — это текст, которого нет в .docx: подписи к иллюстрациям,
    вопросы к ним, врезки. Числа складываются с основными, но помечены
    колонкой «в т.ч. из PDF» и не подсвечены в docx — самого текста подписи
    в документе нет.
    """
    samples: dict[int, dict] = {}
    for c in pdf_result.contexts:
        samples.setdefault(c["unit_idx"], c)

    for idx, count in pdf_result.counts.items():
        if not count:
            continue
        book.counts[idx] += count
        book.pdf[idx] += count
        sample = samples.get(idx)
        book.contexts.append(
            {
                "unit_idx": idx,
                "n": book.counts[idx],
                "variant": sample["variant"] if sample else units[idx].name,
                "kind": "PDF-добавка",
                "found": sample["found"] if sample else "",
                "context": sample["context"] if sample else "",
                "status": "из PDF: подписи и врезки (в docx не подсвечено)",
                "reason": f"вхождений в PDF-добавке: {count}",
            }
        )


def find_books(corpus_root: Path) -> list[Path]:
    """Все учебники-.docx корпуса, кроме явно исключённых дублей."""
    books = [
        p
        for p in sorted(corpus_root.rglob("*.docx"))
        if p.name not in config.SKIP_FILES and not p.name.startswith("~$")
    ]
    return books
