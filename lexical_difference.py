"""
Лексическое обогащение учебников биологии 9, 10 и 11 класса (В. В. Пасечник,
«Линия жизни») — по ТЗ «Алгоритм сравнения лексического обогащения».

Этапы (как в ТЗ):
  1. Очистка: знаки ударения и переносы, номера страниц, колонтитулы, строки
     оглавления, ссылки на рисунки и таблицы, цифры и знаки препинания;
     остаются кириллические слова, дефисные написания и аббревиатуры.
     Разметка основного текста и методического аппарата (рубрики «Вспомните»,
     «Вопросы и задания», «Тренируемся», «Моя лаборатория» и т. п.).
  2. Токенизация с границами предложений, части речи и леммы (Natasha +
     pymorphy3); всё в нижнем регистре, кроме аббревиатур (ДНК, АТФ);
     остаются только NOUN, ADJ, VERB, ADV; вводные слова исключены; имена
     собственные вынесены в отдельный ономастический словарь.
  3. Униграммы F9, F10, F11; биграммы и триграммы внутри предложения по
     моделям ADJ+NOUN, NOUN+NOUN(род. п.), ADJ+ADJ+NOUN (и ADJ+NOUN+NOUN(род. п.),
     NOUN+ADJ(род. п.)+NOUN(род. п.) — как «биологический круговорот веществ»);
     устойчивые — частота ≥ 2 и PMI ≥ 3 или T-score ≥ 2. V9, V10, V11 —
     униграммы + устойчивые словосочетания.
  4. Множества: ядро V9∩V10, V10∩V11; обогащение V10\\V9, V11\\V10;
     выбывшие V9\\V10, V10\\V11.
  5. K_upd, индекс Жаккара, терминологическая плотность D_term; ipm,
     Log-Likelihood G² и Ratio для сквозной лексики (V9∩V10∩V11);
     семантический сдвиг топ-20 сквозных существительных (соседи в окне ±5).
  6. Топ-100 единиц обогащения для экспертной проверки, ономастический словарь.

Запуск: python lexical_difference.py
"""
from __future__ import annotations

import argparse
import json
import math
import re
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from docx import Document
from natasha import Doc
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import config
import highlight_new_and_terms as h
import term_dictionaries as td
import textbook_report as tr
from src.analyzer import normalize_term_to_pattern

GRADES = (9, 10, 11)
OUT = h.HIGHLIGHT_DIR / "Лексическое обогащение_учебники 9, 10 и 11 класса.xlsx"
CACHE_VERSION = 8

CONTENT_POS = {"NOUN", "ADJ", "VERB", "ADV"}
NGRAM_MODELS = {
    ("ADJ", "NOUN"): "ADJ + NOUN",
    ("NOUN", "NOUN_gen"): "NOUN + NOUN (род. п.)",
    ("ADJ", "ADJ", "NOUN"): "ADJ + ADJ + NOUN",
    ("ADJ", "NOUN", "NOUN_gen"): "ADJ + NOUN + NOUN (род. п.)",
    ("NOUN", "ADJ_gen", "NOUN_gen"): "NOUN + ADJ (род. п.) + NOUN (род. п.)",
}
MIN_NGRAM_FREQ = 2
PMI_MIN = 3.0
T_MIN = 2.0
WINDOW = 5
TOP_SHIFT = 20
TOP_CHECK = 100
COMMON_PRIOR = 400     # подсказка «возможно, общеупотребительное»
LL_CRIT = ((15.13, "p < 0,0001"), (10.83, "p < 0,001"), (6.63, "p < 0,01"), (3.84, "p < 0,05"))

# Вводные слова (исключаются как служебные, этап 2)
INTRO_WORDS = {
    "например", "конечно", "впрочем", "итак", "следовательно", "кстати", "наверное", "вероятно",
    "по-видимому", "разумеется", "безусловно", "очевидно", "видимо", "кажется", "напротив", "наоборот",
    "во-первых", "во-вторых", "в-третьих", "в-четвертых", "наконец", "действительно", "возможно",
    "по-моему", "по-нашему", "пожалуй", "так-то", "значит", "словом", "собственно",
}
# Сокращения единиц счёта и измерения — не лексика курса
UNIT_WORDS = {"млн", "млрд", "тыс", "км", "см", "мм", "мкм", "нм", "кг", "мг", "мл", "мин", "сут", "га", "ккал",
              "кдж", "дж", "вт", "т", "г", "л", "м", "с", "ч", "°с"}

# Рубрики методического аппарата (этап 1, сегментация)
PRE_RUBRICS = ("вспомните", "вы узнаете", "вы научитесь")
POST_RUBRICS = ("проверьте себя", "подумайте", "запомните", "моя лаборатория", "исследуйте",
                "выполните задани", "тренируемся", "совершенствуемся", "обсуждаем", "проводим исследование",
                "проекты и исследования", "шаги к успеху", "лабораторный практикум", "лабораторная работа",
                "ход работы", "вопросы и задания")
INFO_RUBRICS = ("это интересно", "для любознательных", "из истории науки", "это важно знать")

FIG_RE = re.compile(r"\(\s*(?:см\.\s*)?(?:рис|табл)\.\s*[\d\s,–—\-а-я]*\)|(?:см\.\s*)?\b(?:[Рр]ис|[Тт]абл)\.\s*\d+[\d,\s–—\-]*(?:[а-я]\b)?")
TOC_LINE_RE = re.compile(r"(?:\.\s?){4,}\s*\d{1,3}\s*$|…\s*\d{1,3}\s*$")
SECTION_RE = re.compile(r"^§\s*\d+")
PAGE_NO_RE = re.compile(r"\d{1,3}")
CHAPTER_RE = re.compile(r"^(?:Глава|ГЛАВА)\s+\d+")
INVISIBLE_RE = re.compile(r"[­​‌‍⁠﻿́´]")
FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
NAME_GRAMS = {"Surn", "Name", "Patr", "Geox", "Orgn"}


def source_path(grade: int) -> Path:
    return Path(config.SOURCE_DOCX[grade] if grade in getattr(config, "SOURCE_DOCX", {}) else config.GRADE_DOCX[grade])


# =============================================================================
# Этап 1. Очистка и сегментация
# =============================================================================

def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.replace("\xa0", " ")).strip().lower().replace("ё", "е")


def toc_titles(grade: int) -> set[str]:
    """Названия параграфов из оглавления PDF (в 10 классе в docx нет заголовков «§ N»)."""
    pdf = getattr(config, "SOURCE_PDF", {}).get(grade)
    if not pdf or not Path(pdf).exists():
        return set()
    try:
        sections, _ = tr.parse_toc(tr.load_pdf(Path(pdf)))
    except StopIteration:
        return set()
    return {_norm(s["title"]) for s in sections if len(s["title"]) > 8}


def clean_paragraphs(grade: int, morph) -> tuple[list[str], list[str], dict]:
    """Тексты абзацев после очистки и метка каждого: main — основной текст, app — аппарат."""
    doc = Document(str(source_path(grade)))
    paras = [p for p in h.iter_paragraphs(doc)
             if not any(a.tag == FALLBACK for a in p._p.iterancestors())]  # запасная копия надписей
    texts = [h.paragraph_run_text(p) for p in paras]
    stripped = [t.strip() for t in texts]
    rep = Counter(s for s in stripped if s)
    titles = toc_titles(grade)
    stats = Counter()

    def title_start(low: str) -> bool:
        # в учебниках, сконвертированных из PDF, название параграфа разбито на строки
        low = low.rstrip(".").strip()
        return bool(low) and (low in titles or (len(low) >= 12 and any(t.startswith(low) for t in titles)))

    labels, mode, pre_block = [], "main", False
    for s in stripped:
        low = _norm(s)
        short = len(s) <= 60
        if (SECTION_RE.match(s) and not TOC_LINE_RE.search(s)) or CHAPTER_RE.match(s) \
                or (short and low.startswith(INFO_RUBRICS)) or title_start(low) \
                or (mode == "app" and not pre_block and PAGE_NO_RE.fullmatch(s)):
            # номер страницы закрывает врезку с вопросами: в учебниках из PDF она стоит
            # посреди параграфа, и дальше продолжается основной текст
            mode, pre_block = "main", False
        elif short and low.startswith(POST_RUBRICS):
            mode, pre_block = "app", False
        elif short and low.startswith(PRE_RUBRICS):
            mode, pre_block = "app", True
        elif pre_block and s and not (s.endswith("?") or len(s) < 120 or re.match(r"^\d+[.)]", s)):
            mode, pre_block = "main", False       # после «Вспомните» вопросы кончились — пошёл текст
        labels.append(mode)

    out = []
    for t, s in zip(texts, stripped):
        letters = re.sub(r"[^A-Za-zА-Яа-яЁё]", "", s)
        if s and TOC_LINE_RE.search(s):
            stats["строк оглавления"] += 1
            out.append("")
            continue
        if s and rep[s] >= 4 and len(letters) >= 2 and len(s) <= 70 and s[0].isupper():
            stats["колонтитулов и повторяющихся заголовков рубрик"] += 1
            out.append("")
            continue
        t2, n = FIG_RE.subn(" ", t)
        stats["ссылок на рисунки и таблицы"] += n
        out.append(INVISIBLE_RE.sub("", t2))
    return out, labels, stats


# =============================================================================
# Этап 2. Токенизация, части речи, леммы
# =============================================================================

JUNK = h.load_not_new_words(config.NOT_NEW_WORDS_PATH)   # мусор распознавания
FUNC_POS = {"PREP", "CONJ", "PRCL", "NPRO", "INTJ", "NUMR"}
# Генетические и химические записи кириллицей: «Аа», «ХУ», «ААВВ», «Са»
NOTATION_RE = re.compile(r"[АаВвСсХхУуЕе]{2,8}")


def _is_abbr(text: str, morph, caps_line: bool) -> bool:
    """Аббревиатура (ДНК, АТФ, ГМО): заглавными, есть в словаре как сокращение или (не в строке,
    набранной заглавными, — там это обычные слова заголовка) неизвестное короткое слово."""
    letters = re.sub(r"[^А-ЯЁа-яё]", "", text)
    if len(letters) < 2 or not letters.isupper() or len(letters) > 7:
        return False
    if any("Abbr" in p.tag for p in morph.parse(text.lower())):
        return True
    return not caps_line and len(letters) <= 5 and not morph.word_is_known(text.lower())


def _caps_word(tok) -> bool:
    letters = re.sub(r"[^А-ЯЁа-яё]", "", tok.text)
    return len(letters) >= 3 and letters.isupper()


def tokenize(text: str, pre, sent_base: int, seg: str) -> list[list]:
    """[текст, лемма, часть речи, падеж, № предложения, сегмент, вид]; вид: w — знаменательное,
    n — имя собственное (кандидат), s — служебное/исключённое, p — знак/цифра (разрыв)."""
    text = pre.replace_letters(pre.remove_accents(text))
    doc = Doc(text)
    doc.segment(pre.segmenter)
    doc.tag_morph(pre.morph_tagger)
    morph = pre.morph
    out, sid = [], sent_base - 1
    for si, sent in enumerate(doc.sents):
        sid = sent_base + si
        toks = sent.tokens
        for k, tok in enumerate(toks):
            t = tok.text
            pos = tok.pos or ""
            case = (tok.feats or {}).get("Case", "")
            if not re.fullmatch(r"[А-Яа-яЁё]+(?:[-‑][А-Яа-яЁё]+)*", t):
                out.append([t, None, pos, case, sid, seg, "p"])
                continue
            low = t.lower()
            letters = re.sub(r"[^а-яё]", "", low)
            known = morph.word_is_known(low)
            notation = NOTATION_RE.fullmatch(t) and (not known or len(set(low)) * 1.5 <= len(low)
                                                     or any(c.isupper() for c in t[1:]))
            fragment = not known and len(letters) <= 5 and re.search(r"(ый|ий|ой|ая|яя|ое|ее)$", low)
            if low in UNIT_WORDS or (len(letters) <= 2 and not known) or notation or fragment or low in JUNK:
                out.append([t, low, pos, case, sid, seg, "s"])        # буквы, записи генотипов, единицы
                continue
            caps_line = (k > 0 and _caps_word(toks[k - 1])) or (k + 1 < len(toks) and _caps_word(toks[k + 1]))
            if _is_abbr(t, morph, caps_line):
                out.append([t, t.upper().replace("Ё", "Е"), "NOUN", case, sid, seg, "w"])
                continue
            tok.lemmatize(pre.morph_vocab)
            lemma = pre.correct_lemma(t, (tok.lemma or t).replace("ё", "е"), []).replace("ё", "е").lower()
            lemma = pre.replacements.get(lemma, lemma)
            if pos == "AUX":
                pos = "VERB"
            kind = "w" if pos in CONTENT_POS else "s"
            best = morph.parse(t)[0]
            grams = set(best.tag.grammemes)
            if re.match(r"^[А-ЯЁ][-‑]", t):
                kind, pos = "w", "NOUN"                           # «Х-хромосома», «В-лимфоцит»
            elif pos == "PROPN":
                if best.tag.POS in FUNC_POS:
                    kind = "s"                                    # «Над», «И» в начале строки
                elif k > 0 or NAME_GRAMS & grams or not known:
                    # не в начале предложения — имя; нарицательные («Земля» и «земля»)
                    # вернёт resolve_names, если лемма где-то пишется как обычное слово
                    kind = "n"
                else:
                    kind, pos = "w", "NOUN"                       # «Опыты» в начале предложения
            elif kind == "w" and t[0].isupper() and k > 0 and pos == "NOUN" and NAME_GRAMS & grams:
                kind = "n"
            if lemma in INTRO_WORDS or lemma in UNIT_WORDS:
                kind = "s"
            if kind == "n" and len(letters) < 3:
                kind = "s"                                        # отдельные буквы «Е», «Д»
            if kind == "n":
                # начальная форма имени — по Natasha без правки pymorphy (тот путает «Линнея» с «линный»)
                lemma = (tok.lemma or low).replace("ё", "е").lower()
            out.append([t, lemma, pos, case, sid, seg, kind])
    return out, sid + 1


def textbook_tokens(grade: int, pre, use_cache: bool) -> tuple[list[list], dict]:
    src = source_path(grade)
    cache = tr.CACHE_DIR / f"lexenrich_{grade}.json"
    if use_cache and cache.exists():
        data = json.loads(cache.read_text(encoding="utf-8"))
        if data.get("version") == CACHE_VERSION and data.get("mtime") == src.stat().st_mtime:
            return data["tokens"], data["stats"]
    texts, labels, stats = clean_paragraphs(grade, pre.morph)
    drop_caps = h.TextStream.find_drop_caps(texts, pre.morph)
    stream = h.TextStream(texts, drop_caps, morph=pre.morph)
    # куски потока с одной меткой (основной текст / аппарат), не длиннее ~100 тыс. символов
    chunks, start, cur = [], 0, None
    for i, o in enumerate(stream.origin):
        lab = labels[o[0]] if o else cur
        if cur is None:
            cur = lab
        if lab != cur or (i - start > 100_000 and stream.text[i] == "\n"):
            chunks.append((start, i, cur))
            start, cur = i, lab
    chunks.append((start, len(stream.text), cur))
    tokens, sid = [], 0
    for a, b, seg in chunks:
        toks, sid = tokenize(stream.text[a:b], pre, sid, seg or "main")
        tokens += toks
    stats = dict(stats, **{"сшито переносов": stream.joins})
    tr.CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps({"version": CACHE_VERSION, "mtime": src.stat().st_mtime,
                                 "tokens": tokens, "stats": stats}, ensure_ascii=False), encoding="utf-8")
    print(f"[Разбор] {grade} класс: {len(tokens)} токенов")
    return tokens, stats


def resolve_names(T: dict) -> set[str]:
    """Имена собственные: «n», чья лемма нигде в трёх учебниках не встречается как нарицательная."""
    common = {t[1] for toks in T.values() for t in toks if t[6] == "w"}
    names = set()
    for toks in T.values():
        for t in toks:
            if t[6] == "n":
                if t[1] in common:
                    t[6] = "w"
                    t[2] = "NOUN" if t[2] == "PROPN" else t[2]
                else:
                    names.add(t[1])
    return names


# =============================================================================
# Этап 3. Униграммы и словосочетания
# =============================================================================

def _role(t) -> str | None:
    if t[6] != "w":
        return None
    if t[2] == "NOUN":
        return "NOUN_gen" if t[3] == "Gen" else "NOUN"
    if t[2] == "ADJ":
        return "ADJ_gen" if t[3] == "Gen" else "ADJ"
    return None


def _fits(roles: tuple, model: tuple) -> bool:
    return all(r == m or (m in ("NOUN", "ADJ") and r == m + "_gen") for r, m in zip(roles, model))


def count_units(tokens: list[list]):
    """Униграммы и n-граммы по моделям; для каждой единицы — сколько вхождений в основном тексте."""
    uni, uni_main = Counter(), Counter()
    ngr, ngr_main, ngr_model, forms = Counter(), Counter(), {}, defaultdict(Counter)
    for t in tokens:
        if t[6] == "w":
            uni[t[1]] += 1
            uni_main[t[1]] += t[5] == "main"
    n = len(tokens)
    for i in range(n):
        for size in (2, 3):
            win = tokens[i:i + size]
            if len(win) < size or len({t[4] for t in win}) != 1:
                continue
            roles = tuple(_role(t) for t in win)
            if None in roles:
                continue
            for model, name in NGRAM_MODELS.items():
                if len(model) == size and _fits(roles, model):
                    key = " ".join(t[1] for t in win)
                    ngr[key] += 1
                    ngr_main[key] += win[0][5] == "main"
                    ngr_model[key] = name
                    forms[key][" ".join(t[0] if t[1].isupper() else t[0].lower() for t in win)] += 1
                    break
    return uni, uni_main, ngr, ngr_main, ngr_model, forms


def association(key: str, f: int, uni: Counter, N: int) -> tuple[float, float]:
    parts = key.split()
    fx = [uni.get(p, 0) for p in parts]
    if f <= 0 or not all(fx):
        return 0.0, 0.0
    expected = math.prod(fx) / N ** (len(parts) - 1)
    pmi = math.log2(f / expected) if expected else 0.0
    t = (f - expected) / math.sqrt(f)
    return round(pmi, 2), round(t, 2)


def display_form(key: str, model: str, forms: Counter, morph) -> str:
    """Словарная форма словосочетания: прилагательное согласуется с главным существительным,
    зависимое в родительном падеже берётся в самой частой форме из текста."""
    surf = forms.most_common(1)[0][0].split() if forms else key.split()
    lem = key.split()
    try:
        if model in ("ADJ + NOUN", "ADJ + ADJ + NOUN", "ADJ + NOUN + NOUN (род. п.)"):
            head_i = 1 if model != "ADJ + ADJ + NOUN" else 2
            head = morph.parse(lem[head_i])[0]
            gender = head.tag.gender or "masc"
            words = []
            for i in range(head_i):
                p = next((x for x in morph.parse(lem[i]) if x.tag.POS in ("ADJF", "PRTF")), None)
                inf = p.inflect({gender, "sing", "nomn"}) if p else None
                words.append(inf.word if inf else lem[i])
            return " ".join(words + [lem[head_i]] + surf[head_i + 1:])
        if model.startswith("NOUN"):
            return " ".join([lem[0]] + surf[1:])
    except Exception:
        pass
    return " ".join(surf)


def build_vocab(tokens, morph):
    uni, uni_main, ngr, ngr_main, ngr_model, forms = count_units(tokens)
    N = sum(uni.values())
    stable = {}
    for key, f in ngr.items():
        if f < MIN_NGRAM_FREQ:
            continue
        pmi, t = association(key, f, uni, N)
        if pmi >= PMI_MIN or t >= T_MIN:
            stable[key] = {"f": f, "main": ngr_main[key], "pmi": pmi, "t": t, "model": ngr_model[key],
                           "form": display_form(key, ngr_model[key], forms[key], morph)}
    freq = dict(uni)
    freq.update({k: v["f"] for k, v in stable.items()})
    main = dict(uni_main)
    main.update({k: v["main"] for k, v in stable.items()})
    return {"uni": uni, "ngr_all": ngr, "stable": stable, "freq": freq, "main": main, "N_content": N,
            "forms": {k: v["form"] for k, v in stable.items()}}


# =============================================================================
# Этап 5. Метрики
# =============================================================================

def log_likelihood(a: int, b: int, n1: int, n2: int) -> float:
    """G² (Rayson, Garside): a, b — частоты в корпусах 1 и 2 объёмом n1, n2; знак — где частота выше
    (плюс — во втором корпусе, т. е. в более старшем классе)."""
    e1 = n1 * (a + b) / (n1 + n2)
    e2 = n2 * (a + b) / (n1 + n2)
    g = 2 * ((a * math.log(a / e1) if a else 0) + (b * math.log(b / e2) if b else 0))
    return round(g if b / n2 >= a / n1 else -g, 2)


def significance(g: float) -> str:
    return next((label for crit, label in LL_CRIT if abs(g) >= crit), "незначимо")


def term_density(tokens, term_patterns: dict[int, set[str]]) -> tuple[int, Counter]:
    """Сколько словоупотреблений покрыто терминами словарей (самое длинное совпадение лемм)."""
    seq = [t[1].lower() for t in tokens if t[6] in ("w", "s", "n") and t[1]]
    maxlen = max(term_patterns) if term_patterns else 1
    i, covered, found = 0, 0, Counter()
    while i < len(seq):
        for n in range(min(maxlen, len(seq) - i), 0, -1):
            key = " ".join(seq[i:i + n])
            if key in term_patterns.get(n, ()):
                covered += n
                found[key] += 1
                i += n
                break
        else:
            i += 1
    return covered, found


def neighbours(tokens, targets: set[str]) -> dict[str, Counter]:
    """Знаменательные соседи в окне ±WINDOW слов внутри предложения."""
    res = defaultdict(Counter)
    words = [t for t in tokens if t[6] in ("w", "s", "n") and t[1]]
    n = len(words)
    for i, t in enumerate(words):
        if t[1] not in targets or t[6] != "w":
            continue
        for j in range(max(0, i - WINDOW), min(n, i + WINDOW + 1)):
            u = words[j]
            if j != i and u[4] == t[4] and u[6] == "w" and u[1] != t[1]:
                res[t[1]][u[1]] += 1
    return res


def cosine(a: Counter, b: Counter) -> float:
    num = sum(a[k] * b[k] for k in set(a) & set(b))
    den = math.sqrt(sum(v * v for v in a.values())) * math.sqrt(sum(v * v for v in b.values()))
    return round(num / den, 3) if den else 0.0


def prior_counts(pre, grades: list[int], use_cache: bool) -> Counter:
    """Упоминания лемм во всех школьных учебниках классов grades (частотные списки и учебники биологии)."""
    per_grade = defaultdict(list)
    for g, cnt in h.rulex_counts(grades, pre).items():
        per_grade[g].append(cnt)
    for g, cnt in h.textbook_counts([g for g in sorted(config.GRADE_DOCX) if g in grades], pre, use_cache).items():
        per_grade[g].append(cnt)
    total = Counter()
    for counters in per_grade.values():
        for w in set().union(*counters):
            total[w] += max(c.get(w, 0) for c in counters)
    return total


def pct(x: float, y: float) -> float:
    return round(x / y * 100, 1) if y else 0.0


# =============================================================================
# Отчёт
# =============================================================================

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--min-freq", type=int, default=config.MIN_KNOWN_FREQ, help="не используется (совместимость)")
    ap.add_argument("--no-cache", action="store_true")
    args = ap.parse_args()
    use_cache = not args.no_cache

    from src.preprocessing import TextPreprocessorWithoutStopwords
    pre = TextPreprocessorWithoutStopwords(stopwords_path=config.STOPWORDS_PATH,
                                           replacements_path=config.REPLACEMENTS_PATH)
    morph = pre.morph
    T, ST = {}, {}
    for g in GRADES:
        T[g], ST[g] = textbook_tokens(g, pre, use_cache)
    names = resolve_names(T)
    V = {g: build_vocab(T[g], morph) for g in GRADES}

    words_all = {g: sum(1 for t in T[g] if t[6] in ("w", "s", "n")) for g in GRADES}
    words_main = {g: sum(1 for t in T[g] if t[6] in ("w", "s", "n") and t[5] == "main") for g in GRADES}
    ipm_base = words_all

    def vset(g):
        return set(V[g]["freq"])

    S = {g: vset(g) for g in GRADES}
    form = {}
    for g in GRADES:
        form.update(V[g]["forms"])
    is_ngram = lambda k: " " in k
    shown = lambda k: form.get(k, k)

    # Термины словарей (общий список для всех классов: словарь терминов + указатели и ФРП 10, 11 кл.)
    term_patterns: dict[int, set[str]] = defaultdict(set)
    for g in (10, 11):
        for src, terms in td.term_lists(g, morph).items():
            for term in terms:
                lemmas, pattern = normalize_term_to_pattern(term, pre)
                if lemmas:
                    term_patterns[len(lemmas)].add(pattern.lower())
    all_terms = set().union(*term_patterns.values())
    dens = {g: term_density(T[g], term_patterns) for g in GRADES}

    before = {g: prior_counts(pre, list(range(2, g)), use_cache) for g in (10, 11)}

    # ---------------- Этап 4: множества -------------------------------------
    zones = [
        ("Лексическое ядро (преемственность) 9–10", "V_core = V9 ∩ V10", S[9] & S[10], "Сквозные понятия курса."),
        ("Лексическое ядро (преемственность) 10–11", "V_core = V10 ∩ V11", S[10] & S[11], "Сквозные понятия курса."),
        ("Лексическое обогащение 10 класса", "V_new = V10 \\ V9", S[10] - S[9], "Новая терминология макроуровня."),
        ("Выбывшая / узкоспециальная лексика 9 класса", "V_dropped = V9 \\ V10", S[9] - S[10],
         "Термины микроуровня, не используемые активно в 10 классе."),
        ("Лексическое обогащение 11 класса", "V_new = V11 \\ V10", S[11] - S[10], "Новая терминология макроуровня."),
        ("Выбывшая / узкоспециальная лексика 10 класса", "V_dropped = V10 \\ V11", S[10] - S[11],
         "Термины микроуровня, не используемые активно в 11 классе."),
    ]
    zone_rows = [[z, f, len(s), sum(1 for k in s if not is_ngram(k)), sum(1 for k in s if is_ngram(k)), i]
                 for z, f, s, i in zones]

    # ---------------- Этап 5: сводные коэффициенты --------------------------
    def kupd(new, old):
        return round(len(S[new] - S[old]) / len(S[new]), 4)

    def jacc(a, b):
        return round(len(S[a] & S[b]) / len(S[a] | S[b]), 4)

    summary = [["Учебники", "Биология, 9, 10 и 11 класс, базовый уровень, под ред. В. В. Пасечника («Линия жизни»)."]]
    for g in GRADES:
        summary += [
            [f"{g} кл.: словоупотреблений (токенов-слов) после очистки", words_all[g]],
            [f"{g} кл.:   из них в основном теоретическом тексте / в методическом аппарате",
             f"{words_main[g]} / {words_all[g] - words_main[g]} ({pct(words_main[g], words_all[g])} % / "
             f"{pct(words_all[g] - words_main[g], words_all[g])} %)"],
            [f"{g} кл.: знаменательных словоупотреблений (NOUN, ADJ, VERB, ADV)", V[g]["N_content"]],
            [f"{g} кл.: |V{g}| — униграмм (лемм) + устойчивых словосочетаний",
             f"{len(S[g])} = {len(V[g]['uni'])} + {len(V[g]['stable'])}"],
            [f"{g} кл.: терминологическая плотность D_term, % (словоупотреблений в составе терминов / все словоупотребления)",
             pct(dens[g][0], words_all[g])],
        ]
    summary += [
        ["K_upd 10 кл. = |V10 \\ V9| / |V10|", kupd(10, 9)],
        ["K_upd 11 кл. = |V11 \\ V10| / |V11|", kupd(11, 10)],
        ["Индекс Жаккара J(9, 10) = |V9 ∩ V10| / |V9 ∪ V10|", jacc(9, 10)],
        ["Индекс Жаккара J(10, 11) = |V10 ∩ V11| / |V10 ∪ V11|", jacc(10, 11)],
        ["Сквозная лексика V9 ∩ V10 ∩ V11 (единиц)", len(S[9] & S[10] & S[11])],
        ["Имён собственных (ономастический словарь, исключены из V)", len(names)],
    ]
    for g in GRADES:
        for k, v in ST[g].items():
            summary.append([f"{g} кл., очистка: {k}", v])

    # ---------------- Списки -------------------------------------------------
    def ipm(g, k):
        return round(V[g]["freq"].get(k, 0) / ipm_base[g] * 1e6, 1)

    def unit_rows(keys, g, extra=None):
        rows = []
        for k in sorted(keys, key=lambda x: (-V[g]["freq"].get(x, 0), x)):
            f = V[g]["freq"].get(k, 0)
            st = V[g]["stable"].get(k, {})
            row = [shown(k), "словосочетание" if is_ngram(k) else "слово", st.get("model", ""), f, ipm(g, k),
                   pct(V[g]["main"].get(k, 0), f), "да" if k.lower() in all_terms else ""]
            if extra:
                row += extra(k)
            rows.append(row)
        return rows

    base_cols = ["Единица", "Тип", "Модель словосочетания", "Частота", "ipm",
                 "В основном тексте, % вхождений", "Есть в словарях терминов"]

    def other(g):
        return lambda k: [V[g]["freq"].get(k, 0) if not is_ngram(k) else V[g]["ngr_all"].get(k, 0)]

    sheets = {}
    sheets["Ядро 9∩10"] = pd.DataFrame(
        [[shown(k), "словосочетание" if is_ngram(k) else "слово", V[9]["freq"][k], V[10]["freq"][k], ipm(9, k), ipm(10, k)]
         for k in sorted(S[9] & S[10], key=lambda x: (-(V[9]["freq"][x] + V[10]["freq"][x]), x))],
        columns=["Единица", "Тип", "Частота 9 кл.", "Частота 10 кл.", "ipm 9 кл.", "ipm 10 кл."])
    sheets["Ядро 10∩11"] = pd.DataFrame(
        [[shown(k), "словосочетание" if is_ngram(k) else "слово", V[10]["freq"][k], V[11]["freq"][k], ipm(10, k), ipm(11, k)]
         for k in sorted(S[10] & S[11], key=lambda x: (-(V[10]["freq"][x] + V[11]["freq"][x]), x))],
        columns=["Единица", "Тип", "Частота 10 кл.", "Частота 11 кл.", "ipm 10 кл.", "ipm 11 кл."])
    sheets["Обогащение 10 (V10∖V9)"] = pd.DataFrame(unit_rows(S[10] - S[9], 10, other(9)),
                                                   columns=base_cols + ["Встречается в 9 кл. (раз, вне V9)"])
    sheets["Выбывшие 9 (V9∖V10)"] = pd.DataFrame(unit_rows(S[9] - S[10], 9, other(10)),
                                                columns=base_cols + ["Встречается в 10 кл. (раз, вне V10)"])
    sheets["Обогащение 11 (V11∖V10)"] = pd.DataFrame(unit_rows(S[11] - S[10], 11, other(10)),
                                                    columns=base_cols + ["Встречается в 10 кл. (раз, вне V10)"])
    sheets["Выбывшие 10 (V10∖V11)"] = pd.DataFrame(unit_rows(S[10] - S[11], 10, other(11)),
                                                  columns=base_cols + ["Встречается в 11 кл. (раз, вне V11)"])

    # Словосочетания (этап 3)
    rows = []
    for g in GRADES:
        for k, st in sorted(V[g]["stable"].items(), key=lambda kv: (-kv[1]["f"], kv[0])):
            rows.append([g, st["form"], st["model"], st["f"], st["pmi"], st["t"], pct(st["main"], st["f"]),
                         "да" if k in all_terms else ""])
    sheets["Словосочетания"] = pd.DataFrame(rows, columns=[
        "Класс", "Словосочетание", "Модель", "Частота", "PMI", "T-score", "В основном тексте, %",
        "Есть в словарях терминов"])

    # Ключевые термины: keyness сквозной лексики + семантический сдвиг (этап 5)
    core3 = S[9] & S[10] & S[11]
    pos_of = defaultdict(Counter)
    for g in GRADES:
        for t in T[g]:
            if t[6] == "w":
                pos_of[t[1]][t[2]] += 1
    noun_lemmas = {l for l, c in pos_of.items() if c.most_common(1)[0][0] == "NOUN"}
    top_nouns = sorted((k for k in core3 if not is_ngram(k) and k in noun_lemmas),
                       key=lambda x: (-sum(V[g]["freq"][x] for g in GRADES), x))[:TOP_SHIFT]
    nb = {g: neighbours(T[g], set(top_nouns)) for g in GRADES}
    shift = {}
    shift_rows = []
    for k in top_nouns:
        tops = {g: [w for w, _ in nb[g][k].most_common(10)] for g in GRADES}
        c910, c1011 = cosine(nb[9][k], nb[10][k]), cosine(nb[10][k], nb[11][k])
        new11 = [w for w in tops[11] if w not in tops[10]]
        shift[k] = (c910, c1011, ", ".join(shown(w) for w in new11))
        shift_rows.append([shown(k)] + [", ".join(f"{shown(w)} ({nb[g][k][w]})" for w in tops[g]) for g in GRADES]
                          + [c910, c1011, round(1 - c1011, 3), ", ".join(shown(w) for w in new11)])
    sheets["Семантический сдвиг"] = pd.DataFrame(shift_rows, columns=[
        "Слово (топ-20 сквозных существительных)", "Соседи в 9 кл. (частота)", "Соседи в 10 кл. (частота)",
        "Соседи в 11 кл. (частота)", "Сходство окружения 9→10 (косинус)", "Сходство окружения 10→11 (косинус)",
        "Семантический сдвиг 10→11 (1 − косинус)", "Новые соседи в 11 кл. (нет в топ-10 10 кл.)"])

    key_rows = []
    for k in core3:
        f9, f10, f11 = (V[g]["freq"][k] for g in GRADES)
        g910 = log_likelihood(f9, f10, ipm_base[9], ipm_base[10])
        g1011 = log_likelihood(f10, f11, ipm_base[10], ipm_base[11])
        i9, i10, i11 = ipm(9, k), ipm(10, k), ipm(11, k)
        sh = shift.get(k, ("", "", ""))
        key_rows.append([shown(k), "словосочетание" if is_ngram(k) else "слово", "да" if k.lower() in all_terms else "",
                         f9, f10, f11, i9, i10, i11, round(i10 / i9, 2) if i9 else "", round(i11 / i10, 2) if i10 else "",
                         g910, significance(g910), g1011, significance(g1011), sh[0], sh[1], sh[2]])
    key_rows.sort(key=lambda r: (-abs(r[13]), r[0]))
    sheets["Ключевые термины (V_core)"] = pd.DataFrame(key_rows, columns=[
        "Единица", "Тип", "Есть в словарях терминов", "Частота 9 кл.", "Частота 10 кл.", "Частота 11 кл.",
        "ipm 9 кл.", "ipm 10 кл.", "ipm 11 кл.", "Ratio ipm10 / ipm9", "Ratio ipm11 / ipm10",
        "G² 9→10 (+ чаще в 10)", "Значимость 9→10", "G² 10→11 (+ чаще в 11)", "Значимость 10→11",
        "Сходство окружения 9→10", "Сходство окружения 10→11", "Новые соседи в 11 кл."])

    # Проверка топ-100 (этап 6)
    check_rows = []
    for g, old in ((10, 9), (11, 10)):
        new = sorted(S[g] - S[old], key=lambda x: (-V[g]["freq"][x], x))[:TOP_CHECK]
        for i, k in enumerate(new, 1):
            prior = before[g].get(k, 0) if not is_ngram(k) else ""
            hint = ""
            if k.lower() in all_terms:
                hint = "термин из словарей"
            elif not is_ngram(k) and prior != "" and prior >= COMMON_PRIOR:
                hint = "возможно, общеупотребительное"
            check_rows.append([f"{g} кл. (V{g}∖V{old})", i, shown(k), "словосочетание" if is_ngram(k) else "слово",
                               V[g]["freq"][k], prior, hint, ""])
    sheets["Проверка топ-100 V_new"] = pd.DataFrame(check_rows, columns=[
        "Список", "№", "Единица", "Тип", "Частота", "Упоминаний во всех школьных учебниках прошлых классов",
        "Подсказка", "Решение эксперта"])

    # Ономастический словарь
    name_cnt = {g: Counter(t[1] for t in T[g] if t[6] == "n") for g in GRADES}
    sheets["Ономастический словарь"] = pd.DataFrame(
        [[n[:1].upper() + n[1:]] + [name_cnt[g][n] for g in GRADES] + [sum(name_cnt[g][n] for g in GRADES)]
         for n in sorted(names, key=lambda x: (-sum(name_cnt[g][x] for g in GRADES), x))],
        columns=["Имя собственное (лемма)", "9 кл.", "10 кл.", "11 кл.", "Всего"])

    guide = [
        ["Выходные данные ТЗ", "1) Словник базовой преемственной лексики — вкладки «Ядро 9∩10», «Ядро 10∩11»; "
                               "2) Словник лексического обогащения 11 класса — «Обогащение 11 (V11∖V10)»; "
                               "3) Ранжированный список ключевых терминов с G² и семантическим сдвигом — "
                               "«Ключевые термины (V_core)» и «Семантический сдвиг»."],
        ["Этап 1. Очистка", "Сняты знаки ударения и мягкие переносы, слова, разорванные переносом на границе строк, "
                            "склеены (дефис составных слов сохраняется); удалены строки оглавления, колонтитулы и "
                            "повторяющиеся заголовки рубрик, ссылки «рис. 45», «табл. 3»; цифры, латиница и знаки "
                            "препинания в подсчёт не входят; отдельные буквы, записи генотипов кириллицей («Аа», «ХУ», «ААВВ») и "
                            "сокращения единиц («млн», «см») исключены. В 10 кл. запасные копии надписей не учитываются дважды."],
        ["Этап 1. Сегментация", "Методический аппарат — текст после рубрик «Вспомните», «Вы узнаете» (до начала "
                                "текста параграфа), «Проверьте себя», «Подумайте», «Запомните», «Моя лаборатория», "
                                "«Тренируемся», «Совершенствуемся», «Обсуждаем», «Проводим исследование», лабораторных "
                                "работ и т. п. — до следующего параграфа, главы или рубрики «Это интересно» (она "
                                "считается основным текстом). Метрики считаются по всему тексту; доля вхождений в "
                                "основном тексте дана в столбце «В основном тексте, %»."],
        ["Этап 2. Леммы", "Natasha + pymorphy3. Существительные — им. п. ед. ч., прилагательные — м. р. ед. ч., "
                          "глаголы — инфинитив. Нижний регистр, кроме аббревиатур (ДНК, АТФ, РНК). Оставлены "
                          "NOUN, ADJ, VERB, ADV; исключены предлоги, союзы, частицы, междометия, местоимения "
                          "(в т. ч. «весь», «каждый»), числа и вводные слова («например», «следовательно» и др.)."],
        ["Этап 2. Имена собственные", "Слова, размеченные как имена собственные (и фамилии с заглавной буквы не в "
                                      "начале предложения), если их лемма ни в одном из трёх учебников не встречается "
                                      "как нарицательное слово. Исключены из V, вынесены в «Ономастический словарь»."],
        ["Этап 3. Словосочетания", f"Скользящее окно внутри предложения по соседним словам (без служебных слов и "
                                   f"знаков между ними), модели: {', '.join(NGRAM_MODELS.values())}. Устойчивые: "
                                   f"частота ≥ {MIN_NGRAM_FREQ} и PMI ≥ {PMI_MIN:g} или T-score ≥ {T_MIN:g}. "
                                   f"PMI = log2(f(xy)·N / (f(x)·f(y))), для триграмм — log2(f(xyz)·N² / (f(x)·f(y)·f(z))); "
                                   f"T-score = (f − ожидаемая частота) / √f; N — число знаменательных словоупотреблений."],
        ["V9, V10, V11", "Все леммы знаменательных слов (униграммы) + устойчивые словосочетания учебника. "
                         "Словосочетание сравнивается с другим учебником как единица его V: если там оно встретилось, "
                         "но не прошло отбор, число вхождений показано в последнем столбце вкладок обогащения."],
        ["K_upd", "Коэффициент лексического обновления: доля единиц словаря учебника, которых нет в словаре "
                  "предыдущего класса."],
        ["Индекс Жаккара", "Пересечение словарей двух учебников / их объединение (1 — словари совпадают)."],
        ["D_term", "Терминологическая плотность: словоупотребления, входящие в термины (словарь терминов, указатели и "
                   "термины ФРП 10 и 11 кл. — один список для всех трёх учебников), к общему числу словоупотреблений."],
        ["ipm", "Частота на миллион словоупотреблений учебника (instances per million)."],
        ["G²", "Log-Likelihood (Rayson, Garside): значимость различия частот в двух учебниках. Плюс — единица "
               "чаще в старшем классе, минус — в младшем. |G²| ≥ 3,84 — p < 0,05; ≥ 6,63 — p < 0,01; "
               "≥ 10,83 — p < 0,001; ≥ 15,13 — p < 0,0001."],
        ["Ratio", "Отношение относительных частот (ipm) в старшем и младшем классе."],
        ["Семантический сдвиг", f"Для {TOP_SHIFT} самых частых сквозных существительных (V9 ∩ V10 ∩ V11) собраны "
                                f"знаменательные соседи в окне ±{WINDOW} слов внутри предложения в каждом учебнике. "
                                f"Сходство окружения — косинус векторов частот соседей (1 — окружение то же); "
                                f"сдвиг = 1 − сходство. Показаны топ-10 соседей и новые соседи 11 кл."],
        ["Проверка топ-100", f"По 100 самых частых единиц обогащения 10 и 11 кл. для экспертной проверки. "
                             "«Возможно, общеупотребительное» — слово не из словарей терминов, которое в школьных "
                             f"учебниках всех предметов прошлых классов встречалось {COMMON_PRIOR} раз и больше. Столбец "
                             "«Решение эксперта» заполняется вручную (выбор из списка)."],
    ]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(OUT), engine="openpyxl") as w:
        pd.DataFrame(summary, columns=["Показатель", "Значение"]).to_excel(w, sheet_name="Сводка", index=False)
        pd.DataFrame(zone_rows, columns=["Зона анализа", "Формула", "Единиц", "Слов (униграмм)", "Словосочетаний",
                                         "Методическая интерпретация"]).to_excel(w, sheet_name="Множества", index=False)
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
        pd.DataFrame(guide, columns=["Что", "Как считается"]).to_excel(w, sheet_name="Пояснения", index=False)

    wb = load_workbook(OUT)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.fill = PatternFill("solid", fgColor="EEF2F5")
        for i, col in enumerate(ws.iter_cols(min_row=1, max_row=min(ws.max_row, 200)), 1):
            longest = max((len(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = max(10, min(60, longest + 2))
        if ws.title in ("Сводка", "Пояснения", "Множества", "Семантический сдвиг"):
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws = wb["Сводка"]; ws.column_dimensions["A"].width = 80; ws.column_dimensions["B"].width = 40
    ws = wb["Пояснения"]; ws.column_dimensions["A"].width = 28; ws.column_dimensions["B"].width = 120
    ws = wb["Проверка топ-100 V_new"]
    dv = DataValidation(type="list", formula1='"термин,общеупотребительное,имя собственное,ошибка разбора"',
                        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H2:H{ws.max_row}")
    wb.save(OUT)

    print(f"[Лексическое обогащение] |V9|={len(S[9])}, |V10|={len(S[10])}, |V11|={len(S[11])}; "
          f"V10∖V9={len(S[10] - S[9])}, V11∖V10={len(S[11] - S[10])}; K_upd 11={kupd(11, 10)}, J(10,11)={jacc(10, 11)}; "
          f"имён собственных {len(names)}")
    print(f"[Лексическое обогащение] Сохранено: {OUT}")


if __name__ == "__main__":
    main()
