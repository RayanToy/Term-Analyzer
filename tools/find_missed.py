# -*- coding: utf-8 -*-
"""Поиск единиц, которые алгоритм мог пропустить в конкретном учебнике.

Проверка полноты. Основной поиск требует, чтобы слова единицы шли подряд
(между ними допускается только пунктуация) и в том же порядке. Здесь то же
самое ищется с послаблениями:

* между словами единицы допускаются посторонние слова;
* порядок слов не важен;
* однословные единицы проверяются на то, что их лемма вообще есть в тексте.

Всё найденное послаблениями, но не попавшее в результат, выводится на
просмотр глазами. Отдельно показываются единицы, которые алгоритм нашёл,
но отфильтровал правилами регистра.

Запуск:  python tools/find_missed.py 06_02_MeTo 09_05_MeTo
"""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from src import analyzer, docio, pipeline
from src.lemmatizers import get_backend
from src.units import build_variants, load_units

#: сколько посторонних слов допускается внутри единицы
EXTRA_WORDS = 3


def scan_book(book, units, index, first_lemmas, backend):
    """Возвращает (результат основного поиска, токены абзацев)."""
    result = pipeline.process_book(book, units, index, first_lemmas, backend)
    views = docio.paragraph_views(docio.open_document(book))
    paragraphs = [backend.tokenize(v.text) if v.text.strip() else [] for v in views]
    return result, paragraphs


def build_lemma_positions(paragraphs):
    """Индекс: лемма -> [(абзац, позиция среди словарных токенов)]."""
    positions = defaultdict(list)
    words_per_para = []
    for para_idx, tokens in enumerate(paragraphs):
        words = [t for t in tokens if t.is_word]
        words_per_para.append(words)
        for i, token in enumerate(words):
            for lemma in token.lemmas:
                positions[lemma].append((para_idx, i))
    return positions, words_per_para


def find_loose(pattern, positions, words_per_para):
    """Ищет все леммы единицы в одном окне, в любом порядке.

    Возвращает контекст первого попадания или None.
    """
    needed = set(pattern)
    window = len(pattern) + EXTRA_WORDS

    # стартуем от самой редкой леммы — так меньше всего проверок
    anchor = min(needed, key=lambda w: len(positions.get(w, ())) or 10**9)
    for para_idx, i in positions.get(anchor, ()):
        words = words_per_para[para_idx]
        lo = max(0, i - window)
        hi = min(len(words), i + window + 1)
        present = set()
        for token in words[lo:hi]:
            present |= token.lemmas
        if needed <= present:
            left = max(0, lo - 4)
            right = min(len(words), hi + 4)
            return " ".join(t.text for t in words[left:right])
    return None


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 1

    units = load_units()
    backend = get_backend(config.PRIMARY_BACKEND)
    build_variants(units, backend)
    index = analyzer.build_pattern_index(units)
    first_lemmas = analyzer.first_lemma_set(index)

    books = [
        b for b in pipeline.find_books(config.CORPUS_ROOT)
        if any(part in b.name for part in argv)
    ]
    if not books:
        print("Учебники не найдены")
        return 1

    for book in books:
        print(f"\n{'=' * 78}\n### {book.parent.name} / {book.stem}\n")
        result, paragraphs = scan_book(book, units, index, first_lemmas, backend)
        positions, words_per_para = build_lemma_positions(paragraphs)

        filtered = []
        loose_hits = []
        single_present = []

        for idx, unit in enumerate(units):
            if result.counts.get(idx, 0):
                continue

            lost = result.rejected.get(idx, 0) + result.lowercase.get(idx, 0)
            if lost:
                filtered.append((lost, unit.name, unit.group_name))
                continue

            for variant in unit.variants:
                if len(variant.pattern) == 1:
                    if variant.pattern[0] in positions:
                        context = find_loose(variant.pattern, positions, words_per_para)
                        single_present.append(
                            (unit.name, unit.group_name, variant.surface, context)
                        )
                        break
                else:
                    context = find_loose(variant.pattern, positions, words_per_para)
                    if context:
                        loose_hits.append(
                            (unit.name, unit.group_name, variant.surface, context)
                        )
                        break

        filtered.sort(reverse=True)
        print(f"--- Есть в тексте, но отсеяно правилами регистра: {len(filtered)} ---")
        for lost, name, group in filtered:
            print(f"  {lost:4}  {name[:42]:44} [{group}]")

        print(f"\n--- Однословные: лемма есть в тексте, но вхождений ноль: "
              f"{len(single_present)} ---")
        for name, group, surface, context in single_present:
            print(f"\n  {name}  [{group}]  вариант {surface!r}")
            print(f"      {(context or '')[:190]}")

        print(f"\n--- Многословные: все слова рядом, но не подряд: "
              f"{len(loose_hits)} ---")
        for name, group, surface, context in loose_hits:
            print(f"\n  {name}  [{group}]  вариант {surface!r}")
            print(f"      {context[:190]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
