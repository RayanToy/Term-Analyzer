# -*- coding: utf-8 -*-
"""Поиск лемма-паттернов в тексте.

Адаптация ``count_patterns_strict`` / ``_words_are_adjacent`` / ``_build_context``
из репозитория Term-Analyzer со следующими изменениями:

* служебные части речи НЕ выбрасываются — иначе рассыпаются пословицы
  («Не имей сто рублей, **а** имей сто друзей») и названия с предлогами
  («окно **в** Европу»);
* словарным считается любой токен с буквой ИЛИ цифрой — нужны «АК-47»,
  «Ту-144», «Восток-1», «58-пушечный»;
* запрет на юниграммы рядом с дефисом снят — он убивал «Ванька-встанька»,
  «Царь-бомба», «Соколов-Микитов»;
* между словами по умолчанию допускается пунктуация («Семь раз отмерь,
  один отрежь»);
* один проход по абзацу даёт и счётчики, и символьные интервалы для
  подсветки — в исходном коде это были два независимых прохода, которые
  могли разойтись.
"""

from __future__ import annotations

from dataclasses import dataclass

import config
from src.lemmatizers import Token

#: пунктуация, после которой заглавная буква обязательна
SENTENCE_END = frozenset(".!?…:;")

#: символы, которые не мешают слову считаться началом предложения
OPENING = frozenset("«\"'([{—–-")

#: сокращения, после точки которых предложение продолжается. Без этого
#: «Историк С. Соловьёв» и «Худ. И. Репин» выглядят как начало предложения,
#: и заглавная буква перестаёт что-либо значить.
#: Список собран по корпусу: взяты частотные короткие слова, стоящие перед
#: точкой, за которой идёт слово с заглавной буквы. Самые частые — «г.»,
#: «в.», «гг.», «реж.», «э.», «вв.», «др.».
ABBREVIATIONS = frozenset({
    "г", "гг", "в", "вв", "им", "худ", "ул", "д", "с", "стр", "т", "тыс",
    "млн", "млрд", "руб", "обл", "р", "кн", "ред", "проф", "акад", "св",
    "др", "пр", "см", "рис", "табл", "ок", "н", "э", "изд", "сост",
    "реж", "сцен", "комп", "авт", "пер", "оп", "ил", "илл", "прим",
    "напр", "ср", "гл", "ч", "кл", "о", "м", "км", "кг", "мин", "сек",
})


@dataclass(slots=True)
class Match:
    """Одно найденное вхождение варианта наименования."""

    unit_idx: int
    variant_idx: int
    para_idx: int
    tok_start: int
    tok_end: int
    char_start: int
    char_end: int
    surface: str
    context: str
    sentence_initial: bool
    status: str = ""   # confident | disputed_accepted | disputed_rejected
    reason: str = ""


class TrieNode:
    """Узел бора лемм."""

    __slots__ = ("children", "payload")

    def __init__(self) -> None:
        self.children: dict[str, "TrieNode"] = {}
        self.payload: list[tuple[int, int]] = []


def build_pattern_index(units) -> TrieNode:
    """Складывает все поисковые паттерны в бор.

    Бор, а не словарь n-грамм, потому что у токена в тексте не одна лемма,
    а множество: pymorphy для «Авроры» самым вероятным разбором даёт «аврор»,
    и поиск по одной основной лемме терял крейсер «Аврора». Бор позволяет
    идти сразу по нескольким веткам без комбинаторного взрыва.
    """
    root = TrieNode()
    for unit_idx, unit in enumerate(units):
        for variant_idx, variant in enumerate(unit.variants):
            node = root
            for lemma in variant.pattern:
                node = node.children.setdefault(lemma, TrieNode())
            node.payload.append((unit_idx, variant_idx))
    return root


def first_lemma_set(root: TrieNode) -> set[str]:
    """Множество первых лемм всех паттернов — для быстрой отсечки позиций."""
    return set(root.children)


def max_pattern_length(root: TrieNode) -> int:
    """Глубина бора."""
    depth = 0
    stack = [(root, 0)]
    while stack:
        node, level = stack.pop()
        depth = max(depth, level)
        for child in node.children.values():
            stack.append((child, level + 1))
    return depth


def _words_are_adjacent(
    word_positions: list[int],
    tokens: list[Token],
    start: int,
    k: int,
    allow_punct_between: bool,
) -> bool:
    """Идут ли k словарных токенов подряд (пунктуация между ними допустима)."""
    for r in range(k - 1):
        idx_a = word_positions[start + r]
        idx_b = word_positions[start + r + 1]
        if idx_b == idx_a + 1:
            continue
        if not allow_punct_between:
            return False
        for t_idx in range(idx_a + 1, idx_b):
            if tokens[t_idx].is_word:
                return False
    return True


def _build_context(tokens: list[Token], tok_start: int, tok_end: int, window: int) -> str:
    """Контекст вокруг совпадения: ±window токенов, склейка через пробел."""
    left = max(0, tok_start - window)
    right = min(len(tokens) - 1, tok_end + window)
    return " ".join(t.text for t in tokens[left : right + 1])


def _is_sentence_initial(tokens: list[Token], tok_idx: int) -> bool:
    """Стоит ли токен в позиции, где заглавная буква обязательна.

    Это начало абзаца или позиция после конечной пунктуации. Прозрачны:
    открывающие кавычки и скобки, а также номера — в учебниках заголовки
    оформлены как «4 Храмы.», и без этого «Храмы» считались бы стоящими
    в середине предложения.
    """
    i = tok_idx - 1
    while i >= 0:
        text = tokens[i].text
        if any(ch.isalpha() for ch in text):
            return False
        if text.isdigit() or all(ch in OPENING for ch in text):
            # номер пункта или открывающая скобка/кавычка — идём дальше влево
            i -= 1
            continue
        if text == "." and _is_abbreviation(tokens, i - 1):
            # инициал или сокращение: «Худ. И. Репин» — не начало предложения
            i -= 2
            continue
        if any(ch in SENTENCE_END for ch in text):
            return True
        return False
    return True


def _is_abbreviation(tokens: list[Token], idx: int) -> bool:
    """Стоит ли перед точкой инициал или общепринятое сокращение."""
    if idx < 0 or not tokens[idx].is_word:
        return False
    text = tokens[idx].text
    if len(text) == 1 and text.isalpha() and text.isupper():
        return True
    return text.lower().replace("ё", "е") in ABBREVIATIONS


def find_matches(
    tokens: list[Token],
    root: TrieNode,
    first_lemmas: set[str],
    para_idx: int,
    allow_punct_between: bool | None = None,
) -> list[Match]:
    """Находит все вхождения паттернов в одном абзаце.

    Обход бора: из каждой стартовой позиции ведём фронт активных узлов и
    расширяем его по всем возможным леммам очередного токена. Между словами
    фразы допускается пунктуация, но не другие слова.
    """
    if allow_punct_between is None:
        allow_punct_between = config.ALLOW_PUNCT_BETWEEN

    word_positions = [i for i, t in enumerate(tokens) if t.is_word]
    n = len(word_positions)
    if not n:
        return []

    lemma_sets = [tokens[i].lemmas for i in word_positions]

    # adjacent[j] — можно ли продолжить фразу с j-1 на j
    adjacent = [True] * n
    for j in range(1, n):
        adjacent[j] = _words_are_adjacent(
            word_positions, tokens, j - 1, 2, allow_punct_between
        )

    matches: list[Match] = []
    seen: set[tuple[int, int, int, int]] = set()

    for i in range(n):
        if not (lemma_sets[i] & first_lemmas):
            continue
        frontier = [(root, i)]
        while frontier:
            next_frontier = []
            for node, j in frontier:
                if j >= n or (j > i and not adjacent[j]):
                    continue
                for lemma in lemma_sets[j]:
                    child = node.children.get(lemma)
                    if child is None:
                        continue
                    if child.payload:
                        _emit(
                            matches, seen, child.payload, tokens,
                            word_positions, i, j, para_idx,
                        )
                    if child.children:
                        next_frontier.append((child, j + 1))
            frontier = next_frontier
    return matches


def _emit(matches, seen, payload, tokens, word_positions, i, j, para_idx) -> None:
    """Добавляет вхождения для одного совпавшего интервала слов [i, j]."""
    tok_start = word_positions[i]
    tok_end = word_positions[j]
    surface = None
    context = None
    initial = None

    for unit_idx, variant_idx in payload:
        key = (unit_idx, variant_idx, tok_start, tok_end)
        if key in seen:
            continue
        seen.add(key)
        if surface is None:
            surface = " ".join(t.text for t in tokens[tok_start : tok_end + 1])
            context = _build_context(tokens, tok_start, tok_end, config.CONTEXT_WINDOW)
            initial = _is_sentence_initial(tokens, tok_start)
        matches.append(
            Match(
                unit_idx=unit_idx,
                variant_idx=variant_idx,
                para_idx=para_idx,
                tok_start=tok_start,
                tok_end=tok_end,
                char_start=tokens[tok_start].start,
                char_end=tokens[tok_end].end,
                surface=surface,
                context=context,
                sentence_initial=initial,
            )
        )


def merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Объединяет пересекающиеся интервалы (для подсветки)."""
    if not spans:
        return []
    spans = sorted(spans)
    merged = [list(spans[0])]
    for start, end in spans[1:]:
        if start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [(a, b) for a, b in merged]
