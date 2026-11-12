"""
Словари терминов для подсветки.

  dict  — общий словарь терминов (data/terms_dictionary.xlsx);
  index — указатель терминов в конце учебника;
  frp   — термины из федеральной рабочей программы (что требуется от ученика).

Термин может входить в несколько словарей; для каждого паттерна (леммы
через пробел) хранится набор словарей. Выделение: цвет слова — по старшему
словарю (ФРП > указатель > общий словарь), подчёркивание — второй словарь
(цветом этого словаря), двойное — если термин во всех трёх.
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import pandas as pd
from docx import Document

import config
from src.analyzer import normalize_term_to_pattern

SOURCES = ("frp", "index", "dict")          # по старшинству
SOURCE_NAMES = {"frp": "ФРП", "index": "указатель учебника", "dict": "словарь терминов"}
# Термин выделяется цветом шрифта (не заливкой): серый, зелёный, жёлтый
# (тёмно-жёлтый — светлый жёлтый на белом не читается)
FONT_COLOR = {"frp": "7F7F7F", "index": "00A04A", "dict": "C49A00"}
COLOR_NAMES = {"frp": "серый", "index": "зелёный", "dict": "жёлтый"}
UNDERLINE_COLOR = FONT_COLOR
NEW_WORD_COLOR = "00A3B8"        # новые слова — бирюзовый

PAGES_RE = re.compile(r"[\s,]*\d{1,3}(?:\s*[,–-]\s*\d{1,3})*\s*$")
LETTER_RE = re.compile(r"^[A-ZА-ЯЁ]$")
# Варианты, добавленные только для поиска («белки-ферменты», «лимитирующие факторы»):
# в списке терминов указателя на вкладке «Словари» не показываются
SEARCH_ONLY: set[str] = set()


def style_for(sources) -> tuple:
    """(цвет шрифта, цвет подчёркивания | None, двойное подчёркивание)."""
    present = [s for s in SOURCES if s in sources]
    if not present:
        return (FONT_COLOR["dict"], None, False)
    color = FONT_COLOR[present[0]]
    if len(present) == 1:
        return (color, None, False)
    # Подчёркивание — цветом второго словаря; во всех трёх — двойное
    return (color, UNDERLINE_COLOR[present[1]], len(present) == 3)


def _with_aliases(text: str) -> list[str]:
    """«Аппарат Гольджи (АГ)» -> ["Аппарат Гольджи", "АГ"]; короткие сокращения отбрасываются."""
    out = []
    main = re.sub(r"\([^)]*\)", " ", text)
    main = re.sub(r"\s+", " ", main).strip(" :;,.")
    if main:
        out.append(main)
    for alias in re.findall(r"\(([^)]*)\)", text):
        alias = alias.strip(" :;,.")
        if len(re.sub(r"[^А-Яа-яЁёA-Za-z]", "", alias)) >= 3:
            out.append(alias)
    return out


def parse_index(path: Path, morph) -> list[str]:
    """
    Указатель терминов: «Термин 27», «Белки:» + «• глобулярные 70»,
    перенос «Биотическое сообщество» / «(биоценоз) 142».
    Подпункт-прилагательное даёт «глобулярные белки» (и «… факторы» для
    «Экологические факторы: • лимитирующие»), подпункт-существительное —
    отдельный термин и «белки-активаторы», многословный — «белки защиты
    и нападения».
    """
    raw = [p.text.replace(" ", " ").strip() for p in Document(str(path)).paragraphs]
    lines = []
    for t in raw:
        if not t or t.upper().startswith("УКАЗАТЕЛЬ") or LETTER_RE.match(t):
            continue
        is_sub = t.startswith(("•", "·", "–", "-"))
        if lines and not is_sub and (t.startswith("(") or t[0].islower()) and not PAGES_RE.search(lines[-1][1]):
            lines[-1] = (lines[-1][0], lines[-1][1] + " " + t)
        else:
            lines.append((is_sub, t))

    terms, head = [], None
    for is_sub, t in lines:
        body = PAGES_RE.sub("", t.lstrip("•·–- ")).strip()
        if not body:
            continue
        if not is_sub:
            head = body.rstrip(":").strip() if body.endswith(":") else None
            terms += _with_aliases(body.rstrip(":"))
            continue
        if not head:
            terms += _with_aliases(body)
            continue
        head_main = _with_aliases(head)[0]
        body = _with_aliases(body)[0]
        words = body.split()
        parses = morph.parse(words[0].lower())
        first = parses[0]
        # «Белки: • движения» — существительное в родительном падеже дополняет
        # главное слово («белки движения»), а не самостоятельный термин
        genitive = any("gent" in p.tag for p in parses[:3])
        if len(words) == 1 and first.tag.POS in ("ADJF", "PRTF"):
            # «Белки: • глобулярные» -> «глобулярные белки»
            terms.append(f"{body} {head_main}")
            # и с главным существительным: «Экологические факторы: • лимитирующие» ->
            # «лимитирующие факторы», «Обмен веществ: • пластический» -> «пластический обмен»
            nouns = [w for w in head_main.split()
                     if any(p.tag.POS == "NOUN" and "nomn" in p.tag for p in morph.parse(w.lower())[:2])]
            if len(head_main.split()) > 1 and nouns:
                terms.append(f"{body} {nouns[0]}")
                SEARCH_ONLY.add(f"{body} {nouns[0]}")
        elif len(words) == 1 and not genitive:
            # «Белки: • ферменты» -> «ферменты» и «белки-ферменты»
            terms.append(body)
            if len(head_main.split()) == 1:
                terms.append(f"{head_main}-{body}")
                SEARCH_ONLY.add(f"{head_main}-{body}")
        elif genitive and first.tag.POS == "NOUN":
            # «Белки: • защиты и нападения» -> «белки защиты и нападения»
            terms.append(f"{head_main} {body}")
        else:
            # «Движущие силы эволюции: • естественный отбор» — самостоятельный термин
            terms.append(body)
    return list(dict.fromkeys(terms))


def parse_frp(path: Path) -> list[str]:
    """Список терминов ФРП: по одному в строке, первая строка — «10 КЛАСС»."""
    terms = []
    for p in Document(str(path)).paragraphs:
        t = p.text.replace(" ", " ").strip()
        if not t or re.fullmatch(r"\d+\s*КЛАСС", t, re.I):
            continue
        terms += _with_aliases(t)
    return list(dict.fromkeys(terms))


def read_dict_xlsx(path: Path) -> list[str]:
    df = pd.read_excel(str(path), header=None, engine="openpyxl")
    terms = df.iloc[:, 0].astype(str).str.strip().replace({"nan": ""})
    return [t for t in dict.fromkeys(terms) if t]


def term_lists(grade: int, morph) -> dict[str, list[str]]:
    lists = {"dict": read_dict_xlsx(config.DICT_PATH)}
    if grade in config.INDEX_DOCX and config.INDEX_DOCX[grade].exists():
        lists["index"] = parse_index(config.INDEX_DOCX[grade], morph)
    if grade in config.FRP_DOCX and config.FRP_DOCX[grade].exists():
        lists["frp"] = parse_frp(config.FRP_DOCX[grade])
    return lists


def load_all_patterns(grade: int, preprocessor):
    """
    Паттерны всех словарей класса.

    Returns:
        (patterns_by_len, verb_shadowed, sources {паттерн: frozenset словарей},
         lists {словарь: [исходные термины]}, term_of {(словарь, паттерн): исходный термин})
    """
    lists = term_lists(grade, preprocessor.morph)
    by_len: dict[int, set[str]] = {}
    verb_shadowed: dict[str, set[str]] = defaultdict(set)
    sources: dict[str, set[str]] = defaultdict(set)
    term_of: dict[tuple[str, str], str] = {}
    for src, terms in lists.items():
        for term in terms:
            lemmas, pattern = normalize_term_to_pattern(term, preprocessor)
            if not lemmas:
                continue
            by_len.setdefault(len(lemmas), set()).add(pattern)
            sources[pattern].add(src)
            term_of.setdefault((src, pattern), term)
            word = term.lower().replace("ё", "е")
            if len(lemmas) == 1 and pattern != word and src == "dict":
                if preprocessor.morph.parse(pattern)[0].tag.POS == "INFN":
                    verb_shadowed[pattern].add(word)
    patterns = {k: sorted(v) for k, v in sorted(by_len.items())}
    counts = {src: sum(1 for s in sources.values() if src in s) for src in lists}
    print(f"[Словари] {grade} класс: " + ", ".join(f"{SOURCE_NAMES[s]} — {len(lists[s])} терминов / {counts[s]} паттернов"
                                                  for s in lists))
    return patterns, dict(verb_shadowed), {k: frozenset(v) for k, v in sources.items()}, lists, term_of
