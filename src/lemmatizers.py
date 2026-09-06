# -*- coding: utf-8 -*-
"""Две реализации морфологического разбора за общим интерфейсом.

Всё остальное в проекте (поиск, подсветка, отчёты) работает со списком
``Token`` и от конкретного бэкенда не зависит.

``PymorphyBackend``
    razdel + pymorphy3, первый (самый вероятный вне контекста) разбор.
    Офлайн, чистый Python, быстрый.

``NatashaBackend``
    Segmenter + NewsMorphTagger из natasha даёт контекстную часть речи,
    которой затем ВЫБИРАЕТСЯ подходящий разбор pymorphy3. Это тот самый
    «двухслойный» подход из Term-Analyzer: natasha снимает грамматическую
    омонимию («стали» — глагол или сталь), pymorphy3 даёт нормальную форму
    и граммемы вроде Geox/Name/Surn.

Граммемы в обоих случаях берутся у pymorphy3 — natasha их в нужном виде
не отдаёт, а правилам разрешения спорных вхождений они необходимы.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache

from src.text_utils import (
    normalize_for_tokenize,
    normalize_lemma,
    remove_accents,
    is_wordlike,
)


@dataclass(slots=True)
class Token:
    """Один токен текста.

    Attributes:
        text: исходный текст токена
        lemma: основная нормализованная лемма (строчная, «ё»→«е»); None для пунктуации
        lemmas: ВСЕ возможные леммы словоформы. Поиск идёт по этому множеству:
            pymorphy для «Авроры» самым вероятным разбором даёт «аврор», и по
            одной только основной лемме крейсер «Аврора» не находился бы
        pos: часть речи в нотации бэкенда
        grammemes: граммемы pymorphy3 (OpenCorpora), для правил по именам собственным
        start: начальная символьная позиция в тексте абзаца
        end: конечная символьная позиция (не включая)
    """

    text: str
    lemma: str | None
    lemmas: frozenset
    pos: str
    grammemes: frozenset
    start: int
    end: int

    @property
    def is_word(self) -> bool:
        return self.lemma is not None


# ---------------------------------------------------------------------------
# pymorphy3 — общий слой для обоих бэкендов
# ---------------------------------------------------------------------------

_MORPH = None


def get_morph():
    """Ленивая инициализация pymorphy3 (создание анализатора дорогое)."""
    global _MORPH
    if _MORPH is None:
        import pymorphy3

        _MORPH = pymorphy3.MorphAnalyzer()
    return _MORPH


@lru_cache(maxsize=400_000)
def _parse_word(word: str, wanted_pos: tuple = ()) -> tuple[str, frozenset, str, frozenset]:
    """Разбор словоформы pymorphy3.

    Returns:
        (основная лемма, все леммы, часть речи, граммемы)

    ``wanted_pos`` — подсказка от контекстного теггера natasha: если среди
    разборов есть подходящий по части речи, основной леммой становится он.

    Кэш по словоформе: в учебнике ~100 тыс. словарных токенов на ~20 тыс.
    уникальных словоформ.
    """
    clean = remove_accents(word)
    parses = get_morph().parse(clean)
    if not parses:
        lemma = normalize_lemma(clean)
        return lemma, frozenset({lemma}), "", frozenset()

    # Все леммы — по ним идёт поиск. Граммемы тоже со всех разборов: имя
    # собственное часто проигрывает нарицательному по вероятности
    # («Орёл» как Geox стоит ниже птицы).
    lemmas = frozenset(normalize_lemma(p.normal_form) for p in parses)
    grammemes = frozenset().union(*(set(p.tag.grammemes) for p in parses))

    best = parses[0]
    if wanted_pos:
        for p in parses:
            if str(p.tag.POS or "") in wanted_pos:
                best = p
                break
    return normalize_lemma(best.normal_form), lemmas, str(best.tag.POS or ""), grammemes


@lru_cache(maxsize=400_000)
def _parse_word_filtered(word: str, wanted_pos: tuple) -> tuple[str, frozenset, str, frozenset]:
    """Разбор, где множество лемм сужено контекстной частью речи.

    Это и есть вклад natasha: в «Стали делать сталь» контекстный теггер
    говорит VERB, и лемма «сталь» из множества уходит. pymorphy без
    контекста оставляет оба варианта.
    """
    clean = remove_accents(word)
    parses = get_morph().parse(clean)
    if not parses:
        lemma = normalize_lemma(clean)
        return lemma, frozenset({lemma}), "", frozenset()

    grammemes = frozenset().union(*(set(p.tag.grammemes) for p in parses))
    matching = [p for p in parses if str(p.tag.POS or "") in wanted_pos] if wanted_pos else []
    chosen = matching or parses

    lemmas = frozenset(normalize_lemma(p.normal_form) for p in chosen)
    best = chosen[0]
    return normalize_lemma(best.normal_form), lemmas, str(best.tag.POS or ""), grammemes


def word_is_proper(word: str) -> bool:
    """Есть ли у словоформы разбор как имени собственного."""
    import config

    _, _, _, grammemes = _parse_word(word)
    return bool(grammemes & config.PROPER_GRAMMEMES)


@lru_cache(maxsize=200_000)
def word_lemmas(word: str) -> frozenset:
    """Все нормализованные леммы словоформы."""
    parses = get_morph().parse(remove_accents(word))
    if not parses:
        return frozenset({normalize_lemma(word)})
    return frozenset(normalize_lemma(p.normal_form) for p in parses)


@lru_cache(maxsize=200_000)
def word_proper_lemmas(word: str) -> frozenset:
    """Леммы, которые словоформа даёт через разборы как имени собственного.

    Позволяет различить два похожих случая:

    * «Авроры» -> {аврор, аврора}: лемма «аврора» приходит из разбора Name,
      значит крейсер здесь уместен;
    * «Соловьёв» -> {соловьёв}: лемма «соловей» приходит только из обычного
      родительного падежа множественного числа, а само слово с заглавной
      буквы — фамилия.

    Русские фамилии массово совпадают с этой формой: «Пирогов» — «пирог»,
    «Воронов» — «ворон», «Волков» — «волк», «Муравьёв» — «муравей».
    """
    import config

    return frozenset(
        normalize_lemma(p.normal_form)
        for p in get_morph().parse(remove_accents(word))
        if set(p.tag.grammemes) & config.PROPER_GRAMMEMES
    )


@lru_cache(maxsize=200_000)
def word_has_common_reading(word: str) -> bool:
    """Есть ли у словоформы разбор БЕЗ пометы имени собственного.

    Нужно, чтобы отличать однозначные имена собственные («Москва» — только
    Geox) от омонимичных («Орёл» — и город, и птица): у вторых помета
    Geox стоит всегда и ничего не доказывает.
    """
    import config

    parses = get_morph().parse(remove_accents(word))
    return any(not (set(p.tag.grammemes) & config.PROPER_GRAMMEMES) for p in parses)


# ---------------------------------------------------------------------------
# Базовый класс
# ---------------------------------------------------------------------------


class LemmatizerBackend:
    """Интерфейс морфологического бэкенда."""

    name = "base"

    def tokenize(self, text: str) -> list[Token]:  # pragma: no cover - интерфейс
        raise NotImplementedError

    def warmup(self) -> None:
        """Прогреть тяжёлые ресурсы до замера времени."""
        self.tokenize("Проверка работы морфологического анализатора.")


# ---------------------------------------------------------------------------
# pymorphy3 + razdel
# ---------------------------------------------------------------------------


class PymorphyBackend(LemmatizerBackend):
    name = "pymorphy"

    def __init__(self) -> None:
        from razdel import tokenize as razdel_tokenize

        self._tokenize = razdel_tokenize

    def tokenize(self, text: str) -> list[Token]:
        text = normalize_for_tokenize(text)
        tokens: list[Token] = []
        for sub in self._tokenize(text):
            raw = sub.text
            if is_wordlike(raw):
                lemma, lemmas, pos, grammemes = _parse_word(raw)
            else:
                lemma, lemmas, pos, grammemes = None, frozenset(), "PUNCT", frozenset()
            tokens.append(
                Token(
                    text=raw,
                    lemma=lemma,
                    lemmas=lemmas,
                    pos=pos,
                    grammemes=grammemes,
                    start=sub.start,
                    end=sub.stop,
                )
            )
        return tokens


# ---------------------------------------------------------------------------
# natasha (контекстный теггер) + pymorphy3 (нормальная форма и граммемы)
# ---------------------------------------------------------------------------

#: соответствие Universal Dependencies -> части речи pymorphy3
UD_TO_PYMORPHY = {
    "NOUN": ("NOUN",),
    "PROPN": ("NOUN",),
    "ADJ": ("ADJF", "ADJS", "PRTF", "PRTS"),
    "VERB": ("VERB", "INFN", "PRTF", "PRTS", "GRND"),
    "AUX": ("VERB", "INFN"),
    "ADV": ("ADVB",),
    "NUM": ("NUMR", "NUMB", "ADJF"),
    "PRON": ("NPRO",),
    "DET": ("ADJF", "NPRO"),
    "ADP": ("PREP",),
    "CCONJ": ("CONJ",),
    "SCONJ": ("CONJ",),
    "PART": ("PRCL",),
    "INTJ": ("INTJ",),
}


class NatashaBackend(LemmatizerBackend):
    name = "natasha"

    def __init__(self) -> None:
        # natasha грузится в обход MorphVocab (он завязан на pymorphy2).
        # Леммы и граммемы везде даёт pymorphy3, natasha — только теггер
        # частей речи. Подробности в src/natasha_loader.py.
        from src.natasha_loader import load

        Doc, Segmenter, NewsEmbedding, NewsMorphTagger = load()

        self._Doc = Doc
        self._segmenter = Segmenter()
        self._morph_tagger = NewsMorphTagger(NewsEmbedding())

    def tokenize(self, text: str) -> list[Token]:
        text = normalize_for_tokenize(text)
        if not text.strip():
            return []

        doc = self._Doc(text)
        doc.segment(self._segmenter)
        doc.tag_morph(self._morph_tagger)

        tokens: list[Token] = []
        for token in doc.tokens:
            raw = token.text
            if is_wordlike(raw):
                wanted = UD_TO_PYMORPHY.get(token.pos or "", ())
                lemma, lemmas, pos, grammemes = _parse_word_filtered(raw, wanted)
                pos = token.pos or pos
            else:
                lemma, lemmas, pos, grammemes = None, frozenset(), "PUNCT", frozenset()
            tokens.append(
                Token(
                    text=raw,
                    lemma=lemma,
                    lemmas=lemmas,
                    pos=pos,
                    grammemes=grammemes,
                    start=token.start,
                    end=token.stop,
                )
            )
        return tokens


# ---------------------------------------------------------------------------

_CACHE: dict[str, LemmatizerBackend] = {}


def get_backend(name: str) -> LemmatizerBackend:
    """Возвращает (и кэширует) бэкенд по имени."""
    if name not in _CACHE:
        if name == "pymorphy":
            _CACHE[name] = PymorphyBackend()
        elif name == "natasha":
            _CACHE[name] = NatashaBackend()
        else:
            raise ValueError(f"Неизвестный бэкенд: {name!r}")
    return _CACHE[name]
