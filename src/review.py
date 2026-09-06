# -*- coding: utf-8 -*-
"""Поиск возможных пропусков многословных единиц.

Основной поиск требует, чтобы слова единицы шли подряд и в том же порядке
(между ними допускается только пунктуация). В учебнике то же название часто
записано иначе:

* вставлено слово — «Софийского собора в **Великом** Новгороде»;
* другой порядок — «в Киев служили Золотые ворота»;
* «стояли **Детинец** Новгородский кремль».

Такие места нельзя автоматически засчитывать: тот же приём находит и явный
мусор («три сестры — мойры» на «Три сестры»). Поэтому они не идут в
счётчики, а собираются на отдельный лист «Возможные пропуски» — для
просмотра глазами.
"""

from __future__ import annotations

from collections import defaultdict

import config


_VERDICTS: dict[tuple[str, str], tuple[str, str]] | None = None


def load_verdicts(path=None) -> dict[tuple[str, str], tuple[str, str]]:
    """Результаты ручной проверки: (наименование, вариант) -> (вердикт, комментарий)."""
    global _VERDICTS
    if _VERDICTS is None:
        path = path or config.REVIEW_VERDICTS_FILE
        out: dict[tuple[str, str], tuple[str, str]] = {}
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = [p.strip() for p in line.split("|")]
                if len(parts) < 3:
                    continue
                name, variant, verdict = parts[0], parts[1], parts[2]
                comment = parts[3] if len(parts) > 3 else ""
                out[(name, variant)] = (verdict, comment)
        _VERDICTS = out
    return _VERDICTS


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


def find_loose(pattern, positions, words_per_para) -> str | None:
    """Ищет все леммы единицы в одном окне, в любом порядке.

    Возвращает контекст первого попадания либо None. Поиск начинается с
    самой редкой леммы, поэтому обходится дёшево.
    """
    needed = set(pattern)
    window = len(pattern) + config.REVIEW_EXTRA_WORDS

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


def collect_possible_misses(units, counts, paragraphs, filtered=()) -> list[dict]:
    """Многословные единицы с нулём вхождений, чьи слова всё же рядом.

    ``filtered`` — единицы, которые в тексте нашлись, но были отсеяны
    правилами регистра. Их сюда брать не нужно: они уже описаны на листе
    «Спорные» с причиной. Иначе «Дворянское гнездо» попадало бы в пропуски,
    хотя словосочетание «дворянских гнёзд» найдено и осознанно отклонено —
    это не роман Тургенева.
    """
    positions, words_per_para = build_lemma_positions(paragraphs)

    rows = []
    for idx, unit in enumerate(units):
        if counts.get(idx, 0) or idx in filtered:
            continue
        for variant in unit.variants:
            if len(variant.pattern) < 2:
                continue
            # У единицы должно быть хотя бы одно редкое слово, иначе поиск
            # «все слова рядом» ловит что угодно: «На дне» находится в
            # «на сегодняшний день», «Новый год» — в «новый вид».
            rarest = min(len(positions.get(w, ())) for w in variant.pattern)
            if rarest > config.REVIEW_MAX_ANCHOR:
                continue
            context = find_loose(variant.pattern, positions, words_per_para)
            if context:
                verdict, comment = load_verdicts().get(
                    (unit.name, variant.surface), ("не проверено", "")
                )
                rows.append(
                    {
                        "unit_idx": idx,
                        "variant": variant.surface,
                        "context": context,
                        "verdict": verdict,
                        "comment": comment,
                    }
                )
                break
    return rows
