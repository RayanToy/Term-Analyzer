# -*- coding: utf-8 -*-
"""Проверка полноты по всему корпусу.

Отвечает на вопрос «что алгоритм мог не найти»:

1. единицы, у которых пустой или подозрительный поисковый паттерн — такие не
   найдутся никогда, и это ошибка разбора списка, а не текста;
2. однословные единицы с нулём вхождений, чья лемма при этом есть в тексте
   (должно быть пусто — иначе поиск теряет прямые совпадения);
3. единицы, отсеянные правилами регистра во всех учебниках сразу;
4. единицы, не найденные нигде, — сверка с сырым текстом по началу слова.

Запуск:  python tools/qa_recall.py
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from src import analyzer, docio, pipeline, review
from src.lemmatizers import get_backend
from src.text_utils import normalize_for_tokenize
from src.units import build_variants, load_units


def norm(text: str) -> str:
    return normalize_for_tokenize(text).replace("ё", "е").replace("Ё", "Е").lower()


def main() -> int:
    units = load_units()
    backend = get_backend(config.PRIMARY_BACKEND)
    build_variants(units, backend)
    index = analyzer.build_pattern_index(units)
    first_lemmas = analyzer.first_lemma_set(index)

    print("=" * 78)
    print("1. Единицы с пустым или вырожденным паттерном")
    broken = [
        u for u in units
        if not u.variants or any(not v.pattern for v in u.variants)
    ]
    print(f"   {len(broken)}")
    for u in broken[:20]:
        print(f"     {u.name!r} [{u.group_sheet}]")

    single_letter = [
        u for u in units
        if u.variants and len(u.variants[0].pattern) == 1
        and len(u.variants[0].pattern[0]) <= 2
    ]
    print(f"   паттерн из одного очень короткого слова: {len(single_letter)}")
    for u in single_letter[:20]:
        print(f"     {u.name!r} -> {u.variants[0].pattern}")

    books = pipeline.find_books(config.CORPUS_ROOT)
    totals: dict[int, int] = defaultdict(int)
    lost: dict[int, int] = defaultdict(int)
    single_present: list = []
    raw_by_book: dict[str, str] = {}

    print()
    print("=" * 78)
    print("2. Однословные: лемма есть в тексте, а вхождений ноль")

    for book in books:
        result = pipeline.process_book(book, units, index, first_lemmas, backend)
        views = docio.paragraph_views(docio.open_document(book))
        raw_by_book[book.stem] = norm(" ".join(v.text for v in views))

        paragraphs = [
            backend.tokenize(v.text) if v.text.strip() else [] for v in views
        ]
        positions, _ = review.build_lemma_positions(paragraphs)

        for idx, unit in enumerate(units):
            count = result.counts.get(idx, 0)
            totals[idx] += count
            lost[idx] += result.rejected.get(idx, 0) + result.lowercase.get(idx, 0)
            if count:
                continue
            for variant in unit.variants:
                if len(variant.pattern) == 1 and variant.pattern[0] in positions:
                    if not (result.rejected.get(idx, 0) or result.lowercase.get(idx, 0)):
                        single_present.append((book.stem, unit.name, variant.surface))
                    break
        print(f"   {book.stem[-30:]}: готово")

    print(f"   найдено случаев: {len(single_present)}")
    for stem, name, surface in single_present[:30]:
        print(f"     {name!r} ({surface}) в {stem[-26:]}")

    print()
    print("=" * 78)
    print("3. Единицы, отсеянные правилами регистра во всех учебниках")
    fully_lost = [
        (lost[idx], units[idx].name, units[idx].group_name)
        for idx in range(len(units))
        if not totals[idx] and lost[idx]
    ]
    fully_lost.sort(reverse=True)
    print(f"   {len(fully_lost)}")
    for n, name, group in fully_lost[:30]:
        print(f"     {n:5}  {name[:40]:42} [{group}]")

    print()
    print("=" * 78)
    print("4. Не найдены нигде: слова единицы всё же стоят в одном окне")
    never = [idx for idx in range(len(units)) if not totals[idx] and not lost[idx]]
    print(f"   единиц без единого вхождения: {len(never)}")

    # Тот же приём, что на листе «Возможные пропуски», но без порога редкости:
    # ищем даже частотные единицы, чтобы ничего не пропустить.
    hits = []
    for book in books:
        views = docio.paragraph_views(docio.open_document(book))
        paragraphs = [
            backend.tokenize(v.text) if v.text.strip() else [] for v in views
        ]
        positions, words_per_para = review.build_lemma_positions(paragraphs)
        for idx in never:
            unit = units[idx]
            for variant in unit.variants:
                if len(variant.pattern) < 2:
                    continue
                context = review.find_loose(variant.pattern, positions, words_per_para)
                if context:
                    hits.append((unit.name, unit.group_name, book.stem, context))
                    break

    seen = set()
    unique = [h for h in hits if not (h[0] in seen or seen.add(h[0]))]
    print(f"   кандидатов: {len(hits)} строк / {len(unique)} наименований")
    print("   (лист «Возможные пропуски» показывает те же случаи с порогом редкости)")
    for name, group, stem, context in unique:
        print()
        print(f"   {name}  [{group}]  — {stem[-26:]}")
        print(f"      {context[:170]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
