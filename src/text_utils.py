# -*- coding: utf-8 -*-
"""Нормализация текста перед токенизацией.

Взято из TextPreprocessor репозитория Term-Analyzer и урезано до того, что
нужно этой задаче:

* снятие диакритики (ударений), но с сохранением «й» и «ё»;
* замена латинских гомоглифов на кириллицу — в списке единиц КК реально
  встречаются смешанные написания («Умa пaлaтa», «Гoл кaк coкoл», «Двa caпoгa
  пaрa»), где часть букв набрана латиницей;
* удаление невидимых символов (мягкий перенос, zero-width);
* нормализация «ё» → «е» для сравнения лемм.

Важно: длина строки при нормализации НЕ меняется — символьные позиции токенов
должны совпадать с позициями в исходном тексте абзаца, иначе поедет подсветка.
"""

import re
import unicodedata

# Невидимые символы, которые ломают токенизацию.
INVISIBLE_RE = re.compile(r"[­​‌‍⁠﻿]")

# Комбинирующие знаки ударения.
COMBINING_ACUTE = "́"
COMBINING_GRAVE = "̀"

# Латинские буквы, визуально неотличимые от кириллических.
LATIN_TO_CYRILLIC = {
    "a": "а", "c": "с", "e": "е", "o": "о", "p": "р", "x": "х", "y": "у",
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М",
    "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У",
}

CYRILLIC_RE = re.compile(r"[А-Яа-яЁё]")
LATIN_RE = re.compile(r"[A-Za-z]")

#: словоподобная последовательность (буквы, цифры, дефисы, апострофы)
WORDLIKE_RE = re.compile(r"[0-9A-Za-zА-Яа-яЁё][0-9A-Za-zА-Яа-яЁё\-'’]*")

#: токен считается словарным, если содержит букву ИЛИ цифру
ALNUM_RE = re.compile(r"[0-9A-Za-zА-Яа-яЁё]")


def remove_invisible(text: str) -> str:
    """Заменяет невидимые символы пробелами (длина строки сохраняется)."""
    return INVISIBLE_RE.sub(" ", text)


def remove_accents(text: str) -> str:
    """Убирает комбинирующие знаки ударения.

    Длина строки может уменьшиться только за счёт самих знаков ударения,
    которых в учебниках практически нет; вызывать эту функцию перед
    вычислением позиций для подсветки нельзя — она используется только
    при нормализации наименований из списка.
    """
    if COMBINING_ACUTE not in text and COMBINING_GRAVE not in text:
        return text
    decomposed = unicodedata.normalize("NFD", text)
    stripped = "".join(
        ch for ch in decomposed if ch not in (COMBINING_ACUTE, COMBINING_GRAVE)
    )
    return unicodedata.normalize("NFC", stripped)


def _fix_word(match: re.Match) -> str:
    """Заменяет латинские гомоглифы внутри одного слова."""
    word = match.group(0)
    if not CYRILLIC_RE.search(word) or not LATIN_RE.search(word):
        # Слово либо чисто кириллическое, либо чисто латинское — не трогаем.
        return word
    return "".join(LATIN_TO_CYRILLIC.get(ch, ch) for ch in word)


def replace_latin_homoglyphs(text: str) -> str:
    """Приводит смешанные кириллично-латинские слова к кириллице.

    Работает только внутри слов, где уже есть хотя бы одна кириллическая
    буква, поэтому английские слова в учебнике остаются нетронутыми.
    Длина строки не меняется.
    """
    return WORDLIKE_RE.sub(_fix_word, text)


def normalize_for_tokenize(text: str) -> str:
    """Нормализация, сохраняющая символьные позиции."""
    return replace_latin_homoglyphs(remove_invisible(text))


def normalize_lemma(lemma: str) -> str:
    """Приводит лемму к каноническому виду для сравнения."""
    return lemma.replace("ё", "е").replace("Ё", "Е").lower()


def is_wordlike(token_text: str) -> bool:
    """Токен считается словарным, если содержит букву или цифру."""
    return bool(ALNUM_RE.search(token_text))
