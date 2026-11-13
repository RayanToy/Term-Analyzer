"""
Двухцветная подсветка учебника: термины + новая лексика класса.

Что делает:
    1. Собирает множество «знакомых» лемм (база сравнения):
        - частотные списки RuLex по всем предметам за 2..N-1 классы
        - учебники биологии предыдущих классов из TEXTBOOKS_DIR
    2. Читает словарь терминов и нормализует его в паттерны лемм.
    3. В целевом учебнике для каждого абзаца вычисляет интервалы:
        - терминов (унигаммы и k-граммы, как в src/highlighter.py)
        - новых слов (лемма не встречалась в предыдущих классах)
    4. Выделяет маркером термины (жёлтый) и новые слова (бирюзовый),
       сохраняя исходное форматирование, цвета шрифта, таблицы и
       изображения (runs режутся, а не пересоздаются, в отличие от
       annotate_docx).
    5. Пишет отчёт xlsx: списки новых слов и найденных терминов с частотами.

Запуск:
    python highlight_new_and_terms.py --grade 11 --in "путь/к/учебнику.docx"

Пример (учебники из Downloads):
    python highlight_new_and_terms.py --grade 10 \
        --in "C:/Users/dom/Downloads/toc_Биология_10_класс_....docx"
"""

from __future__ import annotations

import argparse
import json
import re
import time
import unicodedata
from bisect import bisect_left
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

import config
from config import (
    ALLOW_PUNCT_BETWEEN,
    DICT_PATH,
    DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
    FAMILIAR_WORDS_PATH,
    GRADE_DOCX,
    MIN_KNOWN_FREQ,
    NOT_NEW_WORDS_PATH,
    OUTPUT_DIR,
    REPLACEMENTS_PATH,
    RULEX_DATA_DIR,
    STOPWORDS_PATH,
)
from src.analyzer import (
    _words_are_adjacent,
    build_indexed_words,
    merge_spans,
    normalize_term_to_pattern,
    tokens_with_punct,
)
from src.preprocessing import TextPreprocessorWithoutStopwords

# Цвета выделения — цвет шрифта (по просьбе заказчика, вместо заливки)
COLOR_TERM = ('C49A00', None, False)    # термины без словаря-источника — жёлтый
COLOR_NEW = '00A3B8'                    # новая лексика — бирюзовый

# Белый текст стоит на цветных плашках — его не перекрашиваем
WHITE_FONT = 'FFFFFF'

# Части речи, которые могут считаться новой лексикой
NEW_WORD_POS = {'NOUN', 'ADJ', 'VERB', 'ADV'}

# Минимальная длина леммы для отметки «новое слово»
MIN_NEW_WORD_LEN = 3

# Допустимый вид леммы нового слова (ё уже заменена на е)
CYR_WORD_RE = re.compile(r'^[а-яе]+(?:-[а-яе]+)*$')
LATIN_RE = re.compile(r'[A-Za-z]')

# Переносы из вёрстки PDF: дефис + пробел внутри абзаца («прин- цип»)
# и дефис в конце строки-абзаца («Еди-» / «нообразие»)
WRAP_HYPHENS = '-\u2010\u2011\u00AD'
INLINE_WRAP_RE = re.compile(r'(?<=[А-Яа-яЁё])[-\u2010\u2011][ \t\u00A0]+(?=[а-яё])')
TRAILING_WRAP_RE = re.compile(r'(?<=[А-Яа-яЁё])[-\u2010\u2011\u00AD][ \t\u00A0]*$')
LEADING_WORD_RE = re.compile(r'[ \t\u00A0]*([а-яё]+)')

# После «водо- и …», «одно- или …» дефис висячий, а не перенос
SUSPENDED_HYPHEN_NEXT = {'и', 'или', 'либо', 'а', 'ни'}

# Части, которые пишутся через дефис: «из-за», «кто-то», «кое-что»
HYPHEN_PAIRS = {('из', 'за'), ('из', 'под'), ('все', 'таки')}
HYPHEN_PARTICLES = {'то', 'либо', 'нибудь', 'таки'}
HYPHEN_PREFIX_WORDS = {'кое', 'по'}

# Сколько букв после переноса сверять с написаниями в тексте учебника
VOCAB_PREFIX_TAIL = 3

# Первые части сложных слов, которые пишутся слитно, даже если pymorphy
# знает их как отдельное слово («фото» + «системы» -> «фотосистемы»)
FUSED_PREFIXOIDS = {
    'фото', 'био', 'гидро', 'аэро', 'авто', 'микро', 'макро', 'мега', 'моно',
    'поли', 'нео', 'гео', 'зоо', 'фито', 'термо', 'электро', 'радио', 'теле',
    'метео', 'эко', 'агро', 'морфо', 'цито', 'хемо', 'гемо', 'нейро', 'псевдо',
    'ультра', 'супер', 'гипер', 'анти', 'контр', 'инфра', 'экстра', 'интра',
    'пост', 'пре', 'про', 'эндо', 'экзо', 'мезо', 'палео', 'протео', 'пиро',
}
FUSED_LEFT_POS = {'NUMR', 'PREP', 'CONJ', 'PRCL'}

# Обороты, где слово из словаря терминов употреблено не как термин
NON_TERM_EXPRESSIONS = {('в', 'виде'), ('в', 'виду'), ('в', 'целом')}

# Символы, которые выбрасываются из анализируемого потока: мягкий перенос,
# невидимые пробелы, знаки ударения (комбинирующие и «бо´льший»)
DROP_CHARS = set('\u00AD\u200B\u200C\u200D\u2060\uFEFF\u00B4\u02CA\u02B9')

# Апострофы, которыми в учебниках ставят ударение: «бо'льших», «эмердже'нтность»
STRESS_APOSTROPHES = {"'", "`", chr(0x2019), chr(0x02BC)}
CYR_CHAR_RE = re.compile(r'[А-Яа-яЁё]')

# Абзацы, которые не разрывают перенесённое слово: номер страницы
# («попу-» / «87» / «ляцией») и одиночная буквица («методо-» / «М» / «логию»)
PAGE_NUMBER_RE = re.compile(r'\d{1,6}')
SINGLE_CAP_RE = re.compile(r'[А-ЯЁ]')

# Латиница и цифры: «Pelomyxa», «6XX», «2006» — не русские новые слова
LATIN_OR_DIGIT_RE = re.compile(r'[A-Za-z0-9]')

# Размер блока текста для одной токенизации Natasha
BLOCK_CHARS = 50_000

# Минимальная длина слова «буквица + продолжение» для их склейки
DROP_CAP_MIN_WORD = 4

# Куда складывать результаты
HIGHLIGHT_DIR = OUTPUT_DIR / "highlighted"
CACHE_DIR = OUTPUT_DIR / "cache"

# Содержимое run, которое умеем безопасно пересобрать. Run с чем-то ещё
# (картинки, поля, сноски, символы) не трогаем вовсе.
SAFE_RUN_CHILDREN = {
    qn('w:rPr'), qn('w:t'), qn('w:tab'), qn('w:br'), qn('w:cr'),
    qn('w:noBreakHyphen'), qn('w:softHyphen'), qn('w:lastRenderedPageBreak'),
}


# =============================================================================
# Текст документа
# =============================================================================

def iter_paragraphs(doc: Document):
    """
    Возвращает все абзацы документа, включая таблицы и надписи (textbox).

    Обход идёт по всем элементам w:p в теле документа, поэтому покрывает
    и вложенные структуры, которые doc.paragraphs не отдаёт.
    """
    for p_el in doc.element.body.iter(qn('w:p')):
        yield Paragraph(p_el, doc)


def paragraph_run_text(paragraph: Paragraph) -> str:
    """
    Текст абзаца, склеенный строго из его runs.

    Используется вместо paragraph.text, чтобы символьные позиции токенов
    совпадали с содержимым runs (paragraph.text в новых версиях
    python-docx добавляет текст гиперссылок и сбивает смещения).
    """
    return "".join(r.text for r in paragraph.runs)


def docx_lemmas(path: str | Path, preprocessor) -> Counter:
    """
    Собирает частоты лемм по всему документу (абзацы + таблицы).

    Args:
        path: путь к docx
        preprocessor: экземпляр TextPreprocessor

    Returns:
        Counter {лемма: количество}
    """
    doc = Document(str(path))
    counter: Counter = Counter()

    texts = [p.text for p in iter_paragraphs(doc) if p.text and p.text.strip()]

    # Текст режем на блоки: Natasha работает быстрее на крупных кусках,
    # но держать весь учебник в одном Doc незачем.
    block: list[str] = []
    size = 0
    for t in texts:
        block.append(t)
        size += len(t)
        if size > 100_000:
            counter.update(_block_lemmas("\n".join(block), preprocessor))
            block, size = [], 0
    if block:
        counter.update(_block_lemmas("\n".join(block), preprocessor))

    return counter


def _block_lemmas(text: str, preprocessor) -> list[str]:
    """Леммы словарных токенов одного текстового блока."""
    return [
        t['lemma']
        for t in tokens_with_punct(text, preprocessor)
        if t['lemma']
    ]


def rulex_counts(grades: list[int], preprocessor) -> dict[int, Counter]:
    """
    Частоты слов из списков RuLex (class_N_freq.csv) по классам grades.

    Слова приводятся к той же нормализации, что и леммы текста:
    нижний регистр, ё -> е, словарь замен replacements.txt.

    Args:
        grades: классы (2..11)
        preprocessor: экземпляр TextPreprocessor (нужен словарь замен)

    Returns:
        {класс: Counter {лемма: частота}}
    """
    result: dict[int, Counter] = {}
    for grade in grades:
        path = RULEX_DATA_DIR / f"class_{grade}_freq.csv"
        if not path.exists():
            print(f"[База] ⚠ Список RuLex {grade} класса не найден: {path}")
            continue
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        counter: Counter = Counter()
        for word, freq in zip(df.iloc[:, 0], df.iloc[:, 1]):
            word = word.strip().lower().replace('ё', 'е')
            if not word:
                continue
            try:
                count = float(freq)
            except ValueError:
                count = 1.0
            counter[preprocessor.replacements.get(word, word)] += count
        result[grade] = counter
        print(f"[База] RuLex {grade} класс: {len(counter)} слов")
    return result


def known_lemmas(
    grade: int,
    base: str,
    preprocessor,
    use_cache: bool = True,
    min_freq: int = MIN_KNOWN_FREQ,
) -> tuple[set[str], str]:
    """
    Собирает базу «знакомых» ученику лемм для учебника класса grade.

    Слово знакомо, если за все предыдущие классы оно встретилось
    не меньше min_freq раз. Учебники биологии входят и в корпус RuLex,
    поэтому за каждый класс берётся максимум из двух источников,
    а не сумма — иначе одно упоминание в учебнике считалось бы дважды.

    Args:
        grade: класс размечаемого учебника
        base: 'rulex' | 'textbooks' | 'both'
        preprocessor: экземпляр TextPreprocessor
        use_cache: использовать кэш лемм учебников
        min_freq: сколько упоминаний нужно, чтобы слово считалось знакомым

    Returns:
        (множество лемм, описание базы для легенды)
    """
    parts: list[str] = []
    per_grade: dict[int, list[Counter]] = defaultdict(list)

    rulex_grades = [g for g in range(2, 12) if g < grade]
    if base in ('rulex', 'both') and rulex_grades:
        for g, counter in rulex_counts(rulex_grades, preprocessor).items():
            per_grade[g].append(counter)
        parts.append(
            f"частотные списки {min(rulex_grades)}–{max(rulex_grades)} классов"
        )

    book_grades = [g for g in sorted(GRADE_DOCX) if g < grade]
    if base in ('textbooks', 'both') and book_grades:
        for g, counter in textbook_counts(book_grades, preprocessor, use_cache).items():
            per_grade[g].append(counter)
        parts.append(
            f"учебники биологии {min(book_grades)}–{max(book_grades)} классов"
        )

    total: Counter = Counter()
    for counters in per_grade.values():
        for word in set().union(*counters):
            total[word] += max(c.get(word, 0) for c in counters)

    known = {w for w, n in total.items() if n >= min_freq}
    print(
        f"[База] Лемм в источниках: {len(total)}, "
        f"знакомых (упоминаний: {min_freq} и больше): {len(known)}"
    )
    return known, " и ".join(parts)


def textbook_counts(
    grades: list[int],
    preprocessor,
    use_cache: bool = True,
) -> dict[int, Counter]:
    """
    Частоты лемм в учебниках биологии предыдущих классов.

    Результат по каждому классу кэшируется в output/cache, т.к. разбор
    одного учебника занимает 10-20 секунд.

    Args:
        grades: список классов базы сравнения
        preprocessor: экземпляр TextPreprocessor
        use_cache: читать и писать кэш

    Returns:
        {класс: Counter {лемма: частота}}
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    result: dict[int, Counter] = {}

    for grade in grades:
        cache_path = CACHE_DIR / f"textbook_{grade}_counts.json"
        if use_cache and cache_path.exists():
            result[grade] = Counter(json.loads(cache_path.read_text(encoding='utf-8')))
            print(f"[База] Кэш: {cache_path.name} — {len(result[grade])} лемм")
            continue

        path = GRADE_DOCX[grade]
        if not Path(path).exists():
            print(f"[База] ⚠ Учебник {grade} класса не найден: {path}")
            continue
        t0 = time.time()
        counter = docx_lemmas(path, preprocessor)
        result[grade] = counter
        print(
            f"[База] {grade} класс: {len(counter)} уникальных лемм "
            f"({round(time.time() - t0, 1)} с)"
        )
        if use_cache:
            cache_path.write_text(json.dumps(counter, ensure_ascii=False), encoding='utf-8')

    return result


# =============================================================================
# Словарь терминов
# =============================================================================

def load_patterns(
    dict_path: str | Path,
    preprocessor,
) -> tuple[dict[int, list[str]], dict[str, set[str]]]:
    """
    Читает словарь терминов и переводит его в паттерны лемм по длине.

    Natasha иногда нормализует термин-существительное в глагол
    («род» -> «родиться», «пили» -> «пить»); тогда паттерн совпал бы
    и с «родился», «пьёт». Такие паттерны возвращаются отдельно, чтобы
    при поиске отсеять глагольные формы (см. _is_non_term_usage).

    Args:
        dict_path: путь к xlsx со терминами (один столбец без заголовка)
        preprocessor: экземпляр TextPreprocessor

    Returns:
        ({длина_паттерна: [паттерн, ...]},
         {глагольный паттерн: {исходные термины}})
    """
    df = pd.read_excel(str(dict_path), header=None, engine='openpyxl')
    terms = (
        df.iloc[:, 0]
        .astype(str)
        .str.strip()
        .replace({'nan': ''})
    )
    terms = [t for t in dict.fromkeys(terms) if t]

    patterns_by_len: dict[int, set[str]] = {}
    verb_shadowed: dict[str, set[str]] = defaultdict(set)

    for term in terms:
        lemmas, pattern = normalize_term_to_pattern(term, preprocessor)
        if not lemmas:
            continue
        patterns_by_len.setdefault(len(lemmas), set()).add(pattern)

        word = term.lower().replace('ё', 'е')
        if len(lemmas) == 1 and pattern != word:
            if preprocessor.morph.parse(pattern)[0].tag.POS == 'INFN':
                verb_shadowed[pattern].add(word)

    result = {k: sorted(v) for k, v in sorted(patterns_by_len.items())}
    total = sum(len(v) for v in result.values())
    print(
        f"[Словарь] {len(terms)} терминов → {total} паттернов "
        f"по длинам: { {k: len(v) for k, v in result.items()} }; "
        f"нормализованы в глагол: {len(verb_shadowed)}"
    )
    return result, dict(verb_shadowed)


# =============================================================================
# Сквозной текст документа
# =============================================================================

class TextStream:
    """
    Сквозной текст документа для анализа, со сшитыми переносами.

    Учебники, сконвертированные из PDF, содержат переносы вёрстки:
    «прин- цип» внутри абзаца или «Еди-» / «нообразие», когда каждая
    строка стала отдельным абзацем. Поток склеивает такие слова, не
    меняя сам документ: каждый символ потока помнит, откуда он взят —
    (индекс абзаца, позиция в тексте абзаца). Разделители абзацев ('\\n')
    источника не имеют, а дефисы переносов и знаки ударения в поток
    не попадают.

    Буквица (отдельный абзац из одной заглавной буквы) ставится в начало
    своего абзаца, если он передан в drop_caps.

    Если перенос пришёлся на дефис составного слова («причинно-» /
    «следственные»), дефис сохраняется (см. keep_hyphen). Без morph
    дефис переноса всегда выбрасывается.

    Attributes:
        text: анализируемый текст
        origin: для каждого символа text — (абзац, позиция) или None
        joins: сколько переносов сшито
        kept_hyphens: сколько из них сшито с сохранением дефиса
    """

    def __init__(
        self,
        texts: list[str],
        drop_caps: dict[int, int] | None = None,
        morph=None,
    ):
        chars: list[str] = []
        origin: list[tuple[int, int] | None] = []
        self.joins = 0
        self.kept_hyphens = 0
        self.morph = morph
        self.vocab = _inline_vocab(texts) if morph is not None else Counter()
        self._vocab_sorted = sorted(self.vocab)
        drop_caps = drop_caps or {}
        cap_paras = set(drop_caps.values())

        # Слово продолжается из предыдущего абзаца: ведущие пробелы
        # и пустые абзацы пропускаются без разделителя
        continue_word = False

        def transparent(k: int) -> bool:
            """Абзац, который не должен разрывать перенесённое слово."""
            s = texts[k].strip()
            return k in cap_paras or bool(PAGE_NUMBER_RE.fullmatch(s) or SINGLE_CAP_RE.fullmatch(s))

        for pi, t in enumerate(texts):
            if pi in cap_paras:
                continue  # буквица выводится в начале своего абзаца
            if continue_word and transparent(pi):
                continue  # номер страницы или буквица посреди перенесённого слова
            if pi in drop_caps:
                ci = drop_caps[pi]
                off = len(texts[ci]) - len(texts[ci].lstrip())
                chars.append(texts[ci][off])
                origin.append((ci, off))
                continue_word = True

            skip: set[int] = set()

            for m in INLINE_WRAP_RE.finditer(t):
                nxt = LEADING_WORD_RE.match(t, m.end())
                if nxt and nxt.group(1) in SUSPENDED_HYPHEN_NEXT:
                    continue
                keep = self.keep_hyphen(_word_before(t, m.start()), nxt.group(1) if nxt else '')
                skip.update(range(m.start() + keep, m.end()))
                self.joins += 1
                self.kept_hyphens += keep

            join_next = False
            m = TRAILING_WRAP_RE.search(t)
            if m:
                j = next(
                    (k for k in range(pi + 1, len(texts)) if texts[k].strip() and not transparent(k)),
                    None,
                )
                nxt = LEADING_WORD_RE.match(texts[j]) if j is not None else None
                if nxt and nxt.group(1) not in SUSPENDED_HYPHEN_NEXT:
                    keep = t[m.start()] != '\u00AD' and self.keep_hyphen(
                        _word_before(t, m.start()), nxt.group(1)
                    )
                    join_next = True
                    skip.update(range(m.start() + keep, len(t)))
                    self.joins += 1
                    self.kept_hyphens += keep

            for off, ch in enumerate(t):
                if off in skip:
                    continue
                if continue_word:
                    if ch.isspace():
                        continue
                    continue_word = False
                # Разложенная «й» (и + бреве) — собираем в один символ,
                # иначе NFC в remove_accents сдвинет позиции токенов
                if ch == '\u0306' and chars and chars[-1] in 'иИ':
                    chars[-1] = 'й' if chars[-1] == 'и' else 'Й'
                    continue
                if ch in DROP_CHARS or unicodedata.category(ch) == 'Mn':
                    continue
                if (
                    ch in STRESS_APOSTROPHES
                    and 0 < off < len(t) - 1
                    and CYR_CHAR_RE.match(t[off - 1])
                    and CYR_CHAR_RE.match(t[off + 1])
                ):
                    continue
                chars.append(ch)
                origin.append((pi, off))

            if join_next:
                continue_word = True
            elif not continue_word:
                chars.append('\n')
                origin.append(None)

        self.text = ''.join(chars)
        self.origin = origin

    def keep_hyphen(self, left: str, right: str) -> bool:
        """
        Нужно ли сохранить дефис, на который пришёлся перенос строки.

        Порядок решения:
            1. Устойчивые пары: «из-за», «из-под», «все-таки».
            2. Слитное слово есть в словаре pymorphy — дефис не нужен
               («клет-» + «ках»).
            3. Частицы: «кто-то», «где-либо», «когда-нибудь».
            4. В тексте учебника слово встречается без переноса — берётся
               это написание («причинно-следственные» или «фотосистемы»).
               Сверяется начало слова (перенос + 3 буквы), чтобы
               «кристал-» + «лину» нашлось по «кристаллин».
            5. Аббревиатура слева («АТФ-синтетаза», «РНК-полимераза»).
            6. «по-настоящему», «кое-как».
            7. Обе части — самостоятельные слова («рак-отшельник»,
               «белками-активаторами»), кроме приставочных частей
               («фото-», «микро-») и числительных («двух-»).

        Args:
            left: слово перед дефисом
            right: начало слова после переноса

        Returns:
            True — дефис оставить
        """
        if self.morph is None or not left or not right:
            return False
        l = left.lower().replace('ё', 'е')
        r = right.lower().replace('ё', 'е')
        tail = r[:VOCAB_PREFIX_TAIL]

        if (l, r) in HYPHEN_PAIRS:
            return True
        if self.morph.word_is_known(l + r):
            return False
        if r in HYPHEN_PARTICLES:
            return True
        seen_joined = self._vocab_has_prefix(l + tail)
        if self._vocab_has_prefix(f"{l}-{tail}") and not seen_joined:
            return True
        if seen_joined:
            return False
        if left.isupper() and len(left) >= 2:
            return True
        if l in HYPHEN_PREFIX_WORDS:
            return True
        if l in FUSED_PREFIXOIDS or len(l) < 3 or len(r) < 3:
            return False
        if not (self.morph.word_is_known(l) and self.morph.word_is_known(r)):
            return False
        return self.morph.parse(l)[0].tag.POS not in FUSED_LEFT_POS

    def _vocab_has_prefix(self, prefix: str) -> bool:
        """Есть ли в тексте (внутри строк) слово, начинающееся с prefix."""
        i = bisect_left(self._vocab_sorted, prefix)
        return i < len(self._vocab_sorted) and self._vocab_sorted[i].startswith(prefix)

    @staticmethod
    def find_drop_caps(texts: list[str], morph) -> dict[int, int]:
        """
        Находит буквицы и абзацы, к которым они относятся.

        В сконвертированных учебниках буквица — отдельный абзац из одной
        заглавной буквы, причём не всегда прямо перед своим текстом.
        Абзац-продолжение ищется среди ближайших (±3) по условию: буква +
        первое слово абзаца, начинающегося со строчной, — словарное слово
        («Ф» + «ундаментальные»), а не «Ф» + «чить».

        Args:
            texts: тексты абзацев
            morph: pymorphy3.MorphAnalyzer

        Returns:
            {индекс абзаца с текстом: индекс абзаца с буквицей}
        """
        def continues_wrap(j: int) -> bool:
            """Абзац j продолжает перенос («се-» / «мени») — буквица к нему не относится."""
            for k in range(j - 1, -1, -1):
                s = texts[k].strip()
                if not s or PAGE_NUMBER_RE.fullmatch(s) or SINGLE_CAP_RE.fullmatch(s):
                    continue
                return bool(TRAILING_WRAP_RE.search(texts[k]))
            return False

        result: dict[int, int] = {}
        for i, t in enumerate(texts):
            letter = t.strip()
            if len(letter) != 1 or not re.match(r'[А-ЯЁ]', letter):
                continue
            for d in (1, -1, 2, -2, 3, -3):
                j = i + d
                if j < 0 or j >= len(texts) or j in result or continues_wrap(j):
                    continue
                m = LEADING_WORD_RE.match(texts[j])
                if not m:
                    continue
                word = (letter + m.group(1)).lower()
                # Короткие «По», «Кот», «Ив» — случайные склейки подписей
                # к схемам с предлогами, а не буквицы
                if len(word) >= DROP_CAP_MIN_WORD and morph.word_is_known(word):
                    result[j] = i
                    break
        return result

    def blocks(self, size: int = BLOCK_CHARS):
        """
        Режет поток на блоки примерно по size символов,
        только по границам абзацев.

        Yields:
            (start, end) — интервал блока в координатах потока
        """
        n = len(self.text)
        start = 0
        while start < n:
            end = min(start + size, n)
            if end < n:
                cut = self.text.rfind('\n', start, end)
                if cut <= start:
                    cut = self.text.find('\n', end)
                end = cut + 1 if cut != -1 else n
            yield start, end
            start = end

    def to_paragraphs(self, start: int, end: int) -> list[tuple[int, int, int]]:
        """
        Переводит интервал потока в интервалы абзацев.

        Слово, сшитое из двух абзацев, даёт два интервала. Внутри
        одного абзаца интервал непрерывный (включает дефис и пробел
        переноса, если они были между частями).

        Returns:
            [(индекс абзаца, start, end), ...]
        """
        ranges: dict[int, list[int]] = {}
        for o in self.origin[start:end]:
            if o is None:
                continue
            pi, off = o
            r = ranges.get(pi)
            if r is None:
                ranges[pi] = [off, off + 1]
            else:
                r[0] = min(r[0], off)
                r[1] = max(r[1], off + 1)
        return [(pi, s, e) for pi, (s, e) in ranges.items()]


def _word_before(text: str, pos: int) -> str:
    """Слово (только буквы), стоящее непосредственно перед позицией pos."""
    m = re.search(r'[А-Яа-яЁёA-Za-z]+$', text[:pos])
    return m.group(0) if m else ''


def _inline_vocab(texts: list[str]) -> Counter:
    """
    Частоты написаний слов внутри строк (без переносов) по всему документу,
    включая дефисные («причинно-следственные»). Нужны keep_hyphen, чтобы
    выбрать написание слова, разорванного переносом.
    """
    vocab: Counter = Counter()
    word_re = re.compile(r'[а-яе]+(?:-[а-яе]+)*')
    for t in texts:
        vocab.update(word_re.findall(t.lower().replace('ё', 'е')))
    return vocab


# =============================================================================
# Интервалы подсветки
# =============================================================================

def compute_spans(
    tokens: list[dict],
    patterns_by_len: dict[int, list[str]],
    prev_lemmas: set[str],
    preprocessor,
    stats_terms: Counter,
    stats_new: Counter,
    allow_punct_between: bool = ALLOW_PUNCT_BETWEEN,
    disallow_hyphen_adjacent_for_unigrams: bool = DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
    include_propn: bool = False,
    verb_shadowed: dict[str, set[str]] | None = None,
    not_new: set[str] | None = None,
    stats_skipped: Counter | None = None,
    raw_text: str | None = None,
    stats_derived: dict | None = None,
    occurrences: list | None = None,
    common_lemmas: set[str] | None = None,
    term_sources: dict[str, frozenset] | None = None,
) -> list[tuple[int, int, WD_COLOR_INDEX]]:
    """
    Вычисляет интервалы терминов и новых слов за одну токенизацию.

    Термины имеют приоритет: новое слово, попавшее внутрь термина,
    отдельно не красится.

    Args:
        tokens: токены блока текста из tokens_with_punct()
        patterns_by_len: паттерны терминов
        prev_lemmas: леммы предыдущих классов
        preprocessor: экземпляр TextPreprocessor
        stats_terms: счётчик найденных терминов (мутируется)
        stats_new: счётчик новых слов (мутируется)
        allow_punct_between: допускать пунктуацию внутри фразы
        disallow_hyphen_adjacent_for_unigrams: не красить унигаммы у дефиса
        include_propn: считать ли именами собственные новой лексикой
        verb_shadowed: термины, нормализованные в глагол (см. load_patterns)
        not_new: леммы, которые не выделяются как новые (мусор распознавания
            и редкие, но знакомые ученику слова)
        stats_skipped: счётчик невыделенных из-за not_new (мутируется)
        raw_text: текст блока до нормализации — по нему отсеиваются слова
            на латинице и с цифрами (нормализация превращает «Pelomyxa»
            в кириллицу)
        occurrences: сюда дописываются все найденные вхождения
            (вид 'term' | 'new', термин или лемма, start, end) — для отчёта
            со страницами и расстояниями (мутируется)
        common_lemmas: леммы, которые в учебнике встречаются как
            нарицательные. Слово из этого набора, помеченное как имя
            собственное (заглавная буква в начале предложения: «Шапероны
            связываются…»), проверяется как обычное слово

    Returns:
        Список [(start, end, цвет), ...] в координатах блока,
        отсортированный, без пересечений
    """
    words = build_indexed_words(tokens)
    if not words:
        return []

    # ------------------------------------------------------------------
    # Термины
    # ------------------------------------------------------------------
    term_spans: list[tuple[int, int]] = []
    span_keys: list[tuple[int, int, str]] = []   # для словарей термина

    unis = set(patterns_by_len.get(1, []))
    if unis:
        for w in words:
            key = w['lemma']
            if key not in unis:
                key = _term_by_forms(tokens[w['tok_idx']]['text'], key, unis, preprocessor)
                if key is None:
                    continue
            if disallow_hyphen_adjacent_for_unigrams and (
                w['frag_left'] or w['frag_right'] or w['inner_hyphen']
            ):
                continue
            if _is_non_term_usage(tokens, w, verb_shadowed or {}, preprocessor):
                continue
            # Латинское слово, принятое за кириллицу: кусок «che» из «Escherichia»
            # после замены похожих букв читается как «сне» -> «сон». Смешанные
            # («AТФ» с латинской A — артефакт распознавания) остаются терминами
            if raw_text is not None and _latin_only(raw_text[w['start']:w['end']]) \
                    and not LATIN_RE.search(key):
                continue
            term_spans.append((w['start'], w['end']))
            span_keys.append((w['start'], w['end'], key))
            stats_terms[key] += 1
            if occurrences is not None:
                occurrences.append(('term', key, w['start'], w['end']))

    for k, pats in patterns_by_len.items():
        if k <= 1 or not pats or len(words) < k:
            continue
        patset = set(pats)
        for i in range(len(words) - k + 1):
            if not _words_are_adjacent(words, tokens, i, k, allow_punct_between):
                continue
            phrase = " ".join(w['lemma'] for w in words[i:i + k])
            if phrase not in patset:
                continue
            if raw_text is not None and not LATIN_RE.search(phrase) and any(
                    _latin_only(raw_text[w['start']:w['end']]) for w in words[i:i + k]):
                continue
            term_spans.append((words[i]['start'], words[i + k - 1]['end']))
            span_keys.append((words[i]['start'], words[i + k - 1]['end'], phrase))
            stats_terms[phrase] += 1
            if occurrences is not None:
                occurrences.append(('term', phrase, words[i]['start'], words[i + k - 1]['end']))

    term_spans = merge_spans(term_spans)

    # ------------------------------------------------------------------
    # Новая лексика
    # ------------------------------------------------------------------
    allowed_pos = NEW_WORD_POS | ({'PROPN'} if include_propn else set())
    new_spans: list[tuple[int, int]] = []

    for w in words:
        lemma = w['lemma']
        if lemma in prev_lemmas:
            continue
        if len(lemma) < MIN_NEW_WORD_LEN:
            continue
        # Только кириллические слова: отсекает латынь, цифры и мусор
        # распознавания вроде «территории≫» или висячего «водо-»
        if not CYR_WORD_RE.match(lemma):
            continue
        if raw_text is not None and _mostly_latin_or_digits(raw_text[w['start']:w['end']]):
            continue
        if lemma in preprocessor.russian_stopwords:
            continue
        pos = tokens[w['tok_idx']]['pos']
        if pos == 'PROPN' and common_lemmas and lemma in common_lemmas:
            pos = 'NOUN'
        if pos not in allowed_pos:
            continue
        # Обрывки переносов («полу- чили») — не слова, а артефакты вёрстки.
        # Тире с пробелами («геномика — наука…») обрывком не считается
        if w['frag_left'] or w['frag_right']:
            continue
        if _inside_any(w['start'], w['end'], term_spans):
            continue
        if is_familiar(tokens[w['tok_idx']]['text'], prev_lemmas, preprocessor):
            continue
        # Отобрано вручную: мусор распознавания (data/not_new_words.txt)
        # и редкие, но знакомые слова (data/familiar_words.txt)
        if not_new and lemma in not_new:
            if stats_skipped is not None:
                stats_skipped[lemma] += 1
            if occurrences is not None:
                occurrences.append(('skipped', lemma, w['start'], w['end']))
            continue
        # Образовано от знакомого: «доклеточный», «удойность», «влагоёмкость»
        base = derived_from_familiar(lemma, prev_lemmas, preprocessor)
        if base:
            if stats_derived is not None:
                count, _ = stats_derived.get(lemma, (0, base))
                stats_derived[lemma] = (count + 1, base)
            if occurrences is not None:
                occurrences.append(('derived', lemma, w['start'], w['end']))
            continue
        new_spans.append((w['start'], w['end']))
        stats_new[lemma] += 1
        if occurrences is not None:
            occurrences.append(('new', lemma, w['start'], w['end']))

    new_spans = merge_spans(new_spans)

    if term_sources:
        # Стиль слитого интервала — по всем словарям вошедших в него терминов
        import term_dictionaries
        colored = []
        for ms, me in term_spans:
            srcs = set()
            for s_, e_, key in span_keys:
                if s_ < me and e_ > ms:
                    srcs |= term_sources.get(key, frozenset({'dict'}))
            colored.append((ms, me, term_dictionaries.style_for(srcs)))
    else:
        colored = [(s, e, COLOR_TERM) for s, e in term_spans]
    colored += [(s, e, COLOR_NEW) for s, e in new_spans]
    colored.sort(key=lambda x: x[0])
    return colored


def _term_by_forms(surface: str, lemma: str, unis: set[str], preprocessor) -> str | None:
    """
    Однословный термин по начальной форме pymorphy3, если Natasha дала
    несловарную лемму: «олигосахариды» -> «олигосахарида» вместо
    «олигосахарид», «прионы» -> «приона». Словарные леммы не трогаются,
    чтобы «стали» (стать) не стало «сталью».
    """
    morph = preprocessor.morph
    if morph.word_is_known(lemma):
        return None
    head, sep, last = surface.lower().replace('ё', 'е').rpartition('-')
    for p in morph.parse(last):
        nf = head + sep + p.normal_form.replace('ё', 'е')
        if nf in unis:
            return nf
    return None


def _latin_only(surface: str) -> bool:
    """Латинские буквы есть, кириллических нет: «che», «coli»."""
    return bool(LATIN_RE.search(surface)) and not CYR_CHAR_RE.search(surface)


def _mostly_latin_or_digits(surface: str) -> bool:
    """
    Латинских букв и цифр в слове не меньше, чем кириллицы:
    «Pelomyxa», «6XX», «2006» — да; «a-кристаллин» — нет.
    """
    foreign = len(LATIN_OR_DIGIT_RE.findall(surface))
    return foreign > 0 and foreign >= len(CYR_CHAR_RE.findall(surface))


def _is_non_term_usage(
    tokens: list[dict],
    w: dict,
    verb_shadowed: dict[str, set[str]],
    preprocessor,
) -> bool:
    """
    Слово совпало с однословным термином, но употреблено не как термин.

    - Устойчивые обороты: «в виде», «в виду», «в целом».
    - Термин, нормализованный в глагол («род» -> «родиться»): глагольная
      форма в тексте («родился») — не термин, если она не совпадает
      с термином и не является его словоформой.
    """
    i = w['tok_idx']
    surface = tokens[i]['text'].lower().replace('ё', 'е')
    if i > 0 and (tokens[i - 1]['text'].lower(), surface) in NON_TERM_EXPRESSIONS:
        return True

    term_words = verb_shadowed.get(w['lemma'])
    if not term_words or surface in term_words:
        return False
    parses = preprocessor.morph.parse(surface)
    if any(p.normal_form.replace('ё', 'е') in term_words for p in parses):
        return False
    return parses[0].tag.POS in ('VERB', 'INFN')


def is_familiar(surface: str, known: set[str], preprocessor) -> bool:
    """
    Второе мнение для слова, чьей леммы от Natasha нет в базе.

    Natasha ошибается на редких словах («иона» вместо «ион», «гена»
    вместо «ген»), поэтому проверяются все нормальные формы pymorphy3.
    Составное слово («белок-фермент», «светло-коричневый») знакомо,
    если знакома каждая его часть.

    Args:
        surface: слово в тексте
        known: база знакомых лемм
        preprocessor: экземпляр TextPreprocessor (нужен pymorphy3)

    Returns:
        True, если слово следует считать знакомым
    """
    word = surface.lower().replace('ё', 'е')
    if _forms_known(word, known, preprocessor):
        return True

    parts = [p for p in re.split(r'[-\u2010\u2011]', word) if p]
    return len(parts) > 1 and all(
        p in preprocessor.russian_stopwords or _forms_known(p, known, preprocessor)
        for p in parts
    )


def _forms_known(word: str, known: set[str], preprocessor) -> bool:
    """Есть ли в базе само слово или любая его нормальная форма pymorphy3."""
    if word in known:
        return True
    for parse in preprocessor.morph.parse(word):
        form = parse.normal_form.replace('ё', 'е')
        if preprocessor.replacements.get(form, form) in known:
            return True
    return False


# =============================================================================
# Слова, образованные от знакомых
# =============================================================================

# Приставки и первые части, после отсечения которых остаётся самостоятельное
# слово: «до|клеточный», «не|кодирующий», «супер|самец», «био|печать»
DERIV_PREFIXES = sorted({
    'не', 'до', 'от', 'на', 'над', 'под', 'при', 'пре', 'пере', 'раз', 'рас',
    'без', 'бес', 'вы', 'за', 'по', 'про', 'со', 'из', 'ис', 'об', 'вз', 'воз',
    'вос', 'пред', 'пост', 'меж', 'сверх', 'супер', 'анти', 'контр', 'био',
    'микро', 'макро', 'нано', 'гипер', 'гипо', 'поли', 'моно', 'ультра',
    'дву', 'двух', 'трех', 'четырех', 'много', 'мало', 'само', 'взаимо',
    'полу', 'видео', 'фото', 'авто', 'псевдо', 'внутри', 'около',
}, key=len, reverse=True)

# Суффиксы: окончание -> чем заменить, чтобы получить исходное слово
# («удой|ность» -> «удой», «комбина|ционный» -> «комбинация»)
DERIV_SUFFIXES = [
    ('изованный', ['ь', '']),
    ('изовать', ['ь', '']),
    ('ировать', ['', 'а']),
    ('ированный', ['ировать']),
    ('ительный', ['ить', 'ять']),
    ('ционный', ['ция']),
    ('ованный', ['овать']),
    ('ываться', ['аться', 'ать']),
    ('иваться', ['аться', 'ать']),
    ('оваться', ['', 'а']),
    ('ность', ['ный', 'ной', 'ний', '']),
    ('овать', ['', 'а']),
    ('ывать', ['ать']),
    ('ивать', ['ать']),
    ('ечный', ['ка']),
    ('ость', ['ый', 'ой', 'ий']),
    ('ный', ['', 'а', 'о']),
    ('ый', ['']),
    ('ий', ['']),
]

DERIV_LINKERS = 'ое'   # соединительные гласные: «влаг-о-ёмкость»
DERIV_MIN_STEM = 3     # минимальная длина основы после отсечения суффикса
DERIV_MIN_REST = 4     # минимальная длина слова после приставки
DERIV_MIN_RIGHT = 5    # вторая часть сложного слова: «сома» в «мезосома» — нет
DERIV_MIN_BASE = 4     # итоговое знакомое слово: «ров», «зор», «пир» — слишком коротки
# Короткие корни, из которых действительно складывают слова: «вид-о-специфичность»
DERIV_SHORT_ROOTS = {'вид', 'сух', 'нов', 'нос', 'рог', 'зуб', 'рот', 'лес'}
# Глаголы и причастия: после приставки их можно разбирать дальше, даже если
# промежуточной формы нет в словаре («не|спирализованный», «от|шнуровываться»)
VERBAL_ENDINGS = ('ть', 'ться', 'нный', 'вший', 'щий', 'щийся', 'мый')


def _real_known(word: str, known: set[str], preprocessor) -> bool:
    """
    Основа — настоящее знакомое слово: есть в словаре OpenCorpora и в базе
    знакомых. Отсекает короткие и мусорные леммы частотных списков
    («ров», «онн», «вание»), из-за которых «поровый» разбирался как «по + ровый».
    """
    if len(word) < DERIV_MIN_BASE and word not in DERIV_SHORT_ROOTS:
        return False
    return preprocessor.morph.word_is_known(word) and _forms_known(word, known, preprocessor)


def _stem_known(stem: str, known: set[str], preprocessor) -> str | None:
    """
    Знакомое слово, к которому восходит первая часть сложного слова: та же
    основа с другим окончанием — «влаг» -> «влага», «сух» -> «сухой»,
    «биолог» -> «биолог». Подбор по началу слова («олиг» -> «олигарх»)
    не используется: он даёт случайные совпадения.
    """
    if len(stem) < DERIV_MIN_BASE and stem not in DERIV_SHORT_ROOTS:
        return None
    for end in ('', 'а', 'я', 'ь', 'ый', 'ой', 'ий', 'о', 'е', 'ия'):
        if _real_known(stem + end, known, preprocessor):
            return stem + end
    return None


def _pos_group(word: str, preprocessor) -> str:
    """Часть речи по первому разбору pymorphy3: NOUN, ADJ (с причастиями), VERB или прочее."""
    pos = preprocessor.morph.parse(word)[0].tag.POS or ''
    if pos in ('ADJF', 'ADJS', 'PRTF', 'PRTS'):
        return 'ADJ'
    if pos in ('VERB', 'INFN'):
        return 'VERB'
    return pos


def _is_lemma(word: str, preprocessor) -> bool:
    """Слово стоит в начальной форме: «нос», «протон» — да, «иностранцев» — нет."""
    return any(p.normal_form.replace('ё', 'е') == word for p in preprocessor.morph.parse(word))


def derived_from_familiar(
    word: str,
    known: set[str],
    preprocessor,
    depth: int = 2,
    allow_prefix: bool = True,
) -> str | None:
    """
    Образовано ли слово от знакомого ученику. Если да — возвращает пояснение
    («клеточный», «удой», «влага + емкость»), если нет — None.

    Порядок проверок:
        1. дефисное слово, каждая часть знакома или образована от знакомого:
           «биолого-экологический» -> «биология» + «экологический»;
        2. суффикс: «удойность» -> «удой», «протонный» -> «протон»
           (раньше приставки, иначе «про|тонный», «над|зорный»);
        3. приставка: «доклеточный» -> «клеточный»; разбирать остаток дальше
           можно, только если он словарное слово (иначе «до|декановый» -> «декан»);
        4. два корня через «о/е»: «влагоемкость» -> «влага» + «емкость».
    Получившаяся основа должна быть настоящим словом (словарь OpenCorpora)
    из базы знакомых; правила можно сочетать до depth раз.
    """
    w = word.lower().replace('ё', 'е')
    parts = [p for p in re.split(r'[-\u2010\u2011]', w) if p]

    if len(parts) > 1:
        bases = []
        for p in parts:
            if p in preprocessor.russian_stopwords or _forms_known(p, known, preprocessor):
                bases.append(p)
                continue
            base = _stem_known(p[:-1], known, preprocessor) if p[-1] in DERIV_LINKERS else None
            base = base or (derived_from_familiar(p, known, preprocessor, depth - 1) if depth > 0 else None)
            if not base:
                return None
            bases.append(base)
        return " + ".join(bases)

    for end, replacements in DERIV_SUFFIXES:
        if w.endswith(end) and len(w) - len(end) >= DERIV_MIN_STEM:
            stem = w[:-len(end)]
            for rep in replacements:
                base = stem + rep
                if base == w:
                    continue
                # Без окончания прилагательного должно остаться слово в начальной
                # форме («нос-ый», «протон-ный»), а не форма другого слова
                # («иностранцев-ий» -> «иностранцев» от «иностранец»)
                if rep == '' and end in ('ый', 'ий', 'ный') and not _is_lemma(base, preprocessor):
                    continue
                if _real_known(base, known, preprocessor):
                    return base
                # Разбирать дальше только словарное слово: иначе «додекановый»
                # -> «додеканов» -> «до» + «деканов»
                if depth > 0 and preprocessor.morph.word_is_known(base):
                    sub = derived_from_familiar(base, known, preprocessor, depth - 1, allow_prefix)
                    if sub:
                        return sub
            break  # срабатывает только самое длинное подходящее окончание

    # Приставка отсекается один раз («пост|на|тальный» — уже нет) и не меняет
    # часть речи («до|ломит»: существительное -> глагол — не приставка)
    for prefix in (DERIV_PREFIXES if allow_prefix else ()):
        if w.startswith(prefix) and len(w) - len(prefix) >= DERIV_MIN_REST:
            rest = w[len(prefix):]
            if _real_known(rest, known, preprocessor) and _pos_group(rest, preprocessor) == _pos_group(w, preprocessor):
                return rest
            if depth > 0 and (preprocessor.morph.word_is_known(rest) or rest.endswith(VERBAL_ENDINGS)):
                sub = derived_from_familiar(rest, known, preprocessor, depth - 1, allow_prefix=False)
                if sub:
                    return sub

    for i in range(DERIV_MIN_STEM, len(w) - DERIV_MIN_RIGHT):
        if w[i] not in DERIV_LINKERS:
            continue
        left = _stem_known(w[:i], known, preprocessor)
        if not left:
            continue
        right = w[i + 1:]
        if _real_known(right, known, preprocessor):
            return f"{left} + {right}"
        sub = derived_from_familiar(right, known, preprocessor, depth - 1) if depth > 0 else None
        if sub:
            return f"{left} + {sub}"
    return None


def _inside_any(start: int, end: int, spans: list[tuple[int, int]]) -> bool:
    """Пересекается ли интервал хотя бы с одним из spans."""
    for s, e in spans:
        if start < e and s < end:
            return True
        if s >= end:
            break
    return False


# =============================================================================
# Покраска с сохранением форматирования
# =============================================================================

def apply_colors_to_paragraph(
    paragraph: Paragraph,
    colored: list[tuple[int, int, WD_COLOR_INDEX]],
) -> int:
    """
    Красит указанные интервалы, сохраняя остальное форматирование абзаца.

    Вместо пересборки абзаца каждый затронутый run режется на части
    (с копией его rPr), поэтому жирность, курсив, стили, картинки и
    прочее содержимое остальных runs остаются на месте.

    Runs с картинками и объектами не трогаются.

    Args:
        paragraph: объект Paragraph
        colored: интервалы [(start, end, цвет), ...]

    Returns:
        Количество перекрашенных фрагментов
    """
    if not colored:
        return 0

    painted = 0
    pos = 0

    for run in list(paragraph.runs):
        run_text = run.text
        r_start, r_end = pos, pos + len(run_text)
        pos = r_end

        if not run_text:
            continue
        if any(child.tag not in SAFE_RUN_CHILDREN for child in run._element):
            continue
        # Разрыв страницы тоже даёт '' в run.text — такой run не режем
        if any(br.get(qn('w:type')) == 'page' for br in run._element.iter(qn('w:br'))):
            continue
        if str(run.font.color.rgb) == WHITE_FONT:
            continue

        # Части этого run: (текст, цвет|None)
        pieces: list[tuple[str, WD_COLOR_INDEX | None]] = []
        cursor = r_start
        for s, e, color in colored:
            if e <= r_start or s >= r_end:
                continue
            # cursor не даёт задвоить текст, если интервалы перекрылись
            s, e = max(s, r_start, cursor), min(e, r_end)
            if s >= e:
                continue
            if cursor < s:
                pieces.append((run_text[cursor - r_start:s - r_start], None))
            pieces.append((run_text[s - r_start:e - r_start], color))
            cursor = e
        if not pieces:
            continue
        if cursor < r_end:
            pieces.append((run_text[cursor - r_start:], None))

        _replace_run(run, pieces)
        painted += sum(1 for _, color in pieces if color is not None)

    return painted


NOTE_BIO_COLOR = '1E7B34'   # верхнее число — учебники биологии
NOTE_ALL_COLOR = '1F4E9A'   # нижнее число — все учебники
NOTE_FONT_PT = 5            # размер чисел у термина


# Как показывать числа: 'comments' — примечания Word у каждого термина (читаются
# обычным шрифтом на полосе справа, текст страницы не меняется);
# 'boxes' — мелкие плавающие надписи над и под термином
NOTE_MODE = 'comments'
NOTE_AUTHOR = ''   # без подписи автора у примечаний


NOTE_COMMENT_FONT = ("Arial Narrow", 8)


def comment_style(doc) -> None:
    """Компактный стиль текста примечаний (без него Word берёт крупный обычный стиль
    учебника с отступами, и выноски не помещаются на полосе справа)."""
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn
    from docx.shared import Pt
    s = next((x for x in doc.styles if x.style_id == "CommentText"), None)
    if s is None:
        s = doc.styles.add_style("annotation text", WD_STYLE_TYPE.PARAGRAPH)
        s.element.set(qn("w:styleId"), "CommentText")
    s.font.name, s.font.size = NOTE_COMMENT_FONT[0], Pt(NOTE_COMMENT_FONT[1])
    pf = s.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pf.first_line_indent = pf.left_indent = pf.right_indent = Pt(0)


def add_note_comments(paragraph: Paragraph, notes: list[tuple[int, str, str, int]]) -> int:
    """
    Примечание Word на каждом термине: «Биология: 5:12 7:3 · Все учебники: 2:51 …».
    Раскраска уже разрезала runs по границам терминов, поэтому термин — целые runs.
    """
    document = paragraph.part.document
    bounds, pos = [], 0
    for run in paragraph.runs:
        bounds.append((pos, pos + len(run.text), run))
        pos += len(run.text)
    done = 0
    for off, sup, sub, end, key in notes:
        start = off - 1
        runs = [r for a, b, r in bounds if a < b and a >= start and b <= end]
        if not runs:
            # граница термина внутри run (термин не раскрашен) — берём run с первой буквой
            runs = [r for a, b, r in bounds if a <= start < b][:1]
        if not runs:
            continue
        document.add_comment(runs=[runs[0], runs[-1]], text=f"{key}: био {sup} | все {sub}",
                             author=NOTE_AUTHOR, initials="")
        done += 1
    return done


def add_annotations(paragraph: Paragraph, notes: list[tuple[int, str, str, int]]) -> int:
    """
    Ставит у термина два ряда маленьких чисел: над термином (сколько раз
    он встречался в учебниках биологии прошлых классов) и под ним (во всех
    учебниках прошлых классов).

    Числа — в плавающих надписях без обтекания, привязанных к первой букве
    термина: они не занимают места в строке, поэтому вёрстка учебника
    (переносы строк и страниц) не меняется.

    Позиции считаются по тексту абзаца, который раскраска не меняет
    (runs только режутся), поэтому вызывать можно после apply_colors.
    Пометки ставятся с конца абзаца, чтобы позиции не сдвигались.

    Args:
        paragraph: абзац
        notes: [(позиция после первой буквы термина в тексте абзаца, верхнее, нижнее), ...]

    Returns:
        Сколько пометок поставлено
    """
    if NOTE_MODE == 'comments':
        return add_note_comments(paragraph, notes)
    done = 0
    for off, sup, sub, *_ in sorted(notes, reverse=True):
        pos = 0
        for run in list(paragraph.runs):
            r_start, r_end = pos, pos + len(run.text)
            pos = r_end
            if not (r_start < off <= r_end):
                continue
            if any(child.tag not in SAFE_RUN_CHILDREN for child in run._element):
                break
            target = run
            if off < r_end:
                # Термин кончается внутри run — режем его на две части
                k = off - r_start
                _replace_run(run, [(run.text[:k], None), (run.text[k:], None)])
                # исходный run удалён; нужна часть, которая кончается на off
                target = _run_ending_at(paragraph, off)
                if target is None:
                    break
            anchor = target._element
            for el in reversed(_note_boxes(sup, sub)):
                anchor.addnext(el)
            done += 1
            break
    return done


_VML_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
           'xmlns:w10="urn:schemas-microsoft-com:office:word"')
# Сдвиг надписи от первой буквы термина, пт: влево — на ширину этой буквы,
# по вертикали от верха строки — над строкой и под ней
NOTE_LEFT_PT = -4.0
NOTE_UP_PT = -4.5
NOTE_DOWN_PT = 11.5


_NOTE_IDS = iter(range(1, 10 ** 9))


def _note_box(text: str, color: str, top_pt: float, alt: str):
    """
    Run с плавающей надписью (VML) без обтекания: положение — от символа
    привязки и от верха его строки, в строке места не занимает.
    alt («tn:<номер>:<t|b>») — метка для note_layout, который раздвигает
    наезжающие надписи по координатам из Word.
    """
    from note_layout import text_width
    size = NOTE_FONT_PT
    width = text_width(text, size) + 2
    height = size * 1.25
    text = text.replace('&', '&amp;').replace('<', '&lt;')
    return parse_xml(
        f'<w:r {_VML_NS}><w:rPr><w:noProof/></w:rPr><w:pict>'
        f'<v:rect alt="{alt}" style="position:absolute;margin-left:{NOTE_LEFT_PT}pt;margin-top:{top_pt}pt;'
        f'width:{width:.1f}pt;height:{height:.1f}pt;z-index:251660000;'
        f'mso-position-horizontal-relative:char;mso-position-vertical-relative:line;mso-wrap-style:none" '
        f'filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr>'
        f'<w:spacing w:before="0" w:after="0" w:line="{int(height * 20)}" w:lineRule="exact"/>'
        f'<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:color w:val="{color}"/>'
        f'<w:sz w:val="{int(size * 2)}"/><w:szCs w:val="{int(size * 2)}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:txbxContent></v:textbox>'
        f'<w10:wrap type="none"/></v:rect></w:pict></w:r>')


def _note_boxes(sup: str, sub: str) -> list:
    """Две надписи: над термином — учебники биологии, под ним — все учебники."""
    out = []
    nid = next(_NOTE_IDS)
    if sup:
        out.append(_note_box(sup, NOTE_BIO_COLOR, NOTE_UP_PT, f"tn:{nid}:t"))
    if sub:
        out.append(_note_box(sub, NOTE_ALL_COLOR, NOTE_DOWN_PT, f"tn:{nid}:b"))
    return out


def _run_ending_at(paragraph: Paragraph, off: int) -> Run | None:
    pos = 0
    for run in paragraph.runs:
        pos += len(run.text)
        if pos == off:
            return run
        if pos > off:
            return None
    return None


def _replace_run(run: Run, pieces: list[tuple[str, WD_COLOR_INDEX | None]]) -> None:
    """
    Заменяет run набором новых runs с тем же rPr и заданными цветами.

    Args:
        run: исходный run
        pieces: части [(текст, цвет|None), ...]
    """
    el = run._element
    parent = el.getparent()

    for text, color in pieces:
        if text == '':
            continue
        new_el = deepcopy(el)
        # Оставляем только форматирование, содержимое задаём сами
        for child in list(new_el):
            if child.tag != qn('w:rPr'):
                new_el.remove(child)
        _append_run_content(new_el, text)

        el.addprevious(new_el)
        if color is not None:
            _apply_style(Run(new_el, run._parent), color)

    parent.remove(el)


def _keep_color(val: str | None) -> bool:
    """Цвет обычного текста: почти чёрный (231F20 и т. п.) или белый (надписи на плашках)."""
    if not val or val.lower() == "auto" or len(val) != 6:
        return True
    try:
        r, g, b = (int(val[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return True
    return max(r, g, b) < 0x60 or min(r, g, b) > 0xE8


def strip_publisher_colors(doc) -> int:
    """Снимает цветное выделение издателя (фиолетовые, зелёные, синие слова), чтобы
    в тексте остались только наши цвета: текст становится чёрным, белый не трогается."""
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    roots = [doc.element.body, doc.styles.element]
    done = 0
    # цвет обычного текста учебника (в 10 классе 231F20, в 11-м чёрный) — им и заменяем
    from collections import Counter
    dark = Counter()
    for el in doc.element.body.iter(w + "color"):
        v = el.get(w + "val") or ""
        if len(v) == 6 and v.lower() != "auto" and max(int(v[i:i + 2], 16) for i in (0, 2, 4)) < 0x60:
            dark[v.upper()] += 1
    body = dark.most_common(1)[0][0] if dark else "000000"
    for root in roots:
        for el in list(root.iter(w + "color")):
            if _keep_color(el.get(w + "val")):
                continue
            el.set(w + "val", body)
            for attr in ("themeColor", "themeTint", "themeShade"):
                el.attrib.pop(w + attr, None)
            done += 1
        for el in list(root.iter(w + "highlight")):
            el.getparent().remove(el)
            done += 1
    return done


def _apply_style(run: Run, style) -> None:
    """Цвет шрифта (строка RGB) или стиль словарей (цвет шрифта, цвет подчёркивания | None, двойное)."""
    if not isinstance(style, tuple):
        style = (style, None, False)
    color, ucolor, double = style
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if ucolor:
        rpr = run._element.get_or_add_rPr()
        for old in rpr.findall(qn('w:u')):
            rpr.remove(old)
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'double' if double else 'thick')
        u.set(qn('w:color'), ucolor)
        rpr.append(u)


def _append_run_content(run_el, text: str) -> None:
    """
    Записывает текст в run, возвращая табуляции и переносы строк
    в виде w:tab / w:br (run.text отдаёт их как '\\t' и '\\n').
    """
    buf = ''
    for ch in text:
        if ch in '\t\n':
            if buf:
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = buf
                run_el.append(t)
                buf = ''
            run_el.append(OxmlElement('w:tab' if ch == '\t' else 'w:br'))
        else:
            buf += ch
    if buf:
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = buf
        run_el.append(t)


# =============================================================================
# Отчёт
# =============================================================================

DICT_COLUMN = 'Есть в словаре OpenCorpora (pymorphy3)'


def dictionary_note(morph) -> str:
    """Пояснение к колонке «Есть в словаре…»: что за словарь и как читать пустую ячейку."""
    meta = morph.dictionary.meta
    lexemes = f"{meta.get('source_lexemes_count', 0):,}".replace(',', ' ')
    return (
        f"морфологический словарь русского языка OpenCorpora (версия {meta.get('source_version')}, "
        f"{lexemes} слов), которым пользуется библиотека pymorphy3. Пусто — слова в нём нет: "
        f"это редкий научный термин или слово, испорченное при конвертации PDF в docx"
    )


def save_report(
    out_path: Path,
    stats_new: Counter,
    stats_terms: Counter,
    patterns_total: int,
    known: set[str],
    new_terms: set[str],
    sources: str,
    min_freq: int,
    morph,
    stats_skipped: Counter,
    junk: set[str],
    stats_derived: dict,
) -> None:
    """
    Сохраняет списки новых слов и найденных терминов в xlsx.

    new_terms — термины, незнакомые ученику (в документе они выделены
    как термины, маркером новых слов не дублируются).

    stats_skipped — невыделенные по спискам исключений; junk отличает
    мусор распознавания от редких, но знакомых слов.

    Колонка «Есть в словаре OpenCorpora (pymorphy3)» помогает при ручной
    проверке: слова, которых нет в этом словаре, — это либо редкие научные
    термины, либо мусор распознавания из сконвертированного PDF.
    """
    out_path.parent.mkdir(parents=True, exist_ok=True)
    new_df = pd.DataFrame(
        [
            (word, count, 'да' if morph.word_is_known(word) else '')
            for word, count in sorted(stats_new.items(), key=lambda x: (-x[1], x[0]))
        ],
        columns=['Новое слово (лемма)', 'Вхождений в учебнике', DICT_COLUMN],
    )
    terms_df = pd.DataFrame(
        [
            (term, count, 'да' if term in new_terms else '')
            for term, count in sorted(stats_terms.items(), key=lambda x: (-x[1], x[0]))
        ],
        columns=['Термин (слова в начальной форме)', 'Вхождений в учебнике', 'Есть незнакомое ученику слово'],
    )
    summary_df = pd.DataFrame(
        [
            [
                'Слово считается новым, если',
                f'в предыдущих классах оно встретилось меньше {min_freq} раз'
                if min_freq > 1 else 'в предыдущих классах его нет',
            ],
            ['Где считались упоминания', sources],
            [f'Знакомых слов (упоминаний: {min_freq} и больше)', len(known)],
            ['Новых слов (разных)', len(stats_new)],
            ['Новых слов (всего вхождений в учебнике)', sum(stats_new.values())],
            ['Терминов найдено (разных)', len(stats_terms)],
            ['  из них с незнакомым ученику словом', int((terms_df['Есть незнакомое ученику слово'] == 'да').sum())],
            ['Терминов найдено (всего вхождений в учебнике)', sum(stats_terms.values())],
            ['Терминов в словаре (после приведения к начальной форме)', patterns_total],
            ['Не выделено: мусор распознавания (лемм)', sum(1 for w in stats_skipped if w in junk)],
            ['Не выделено: редкие, но знакомые слова (лемм)', sum(1 for w in stats_skipped if w not in junk)],
            ['Не выделено: образованы от знакомых слов (лемм)', len(stats_derived)],
            [f'Что такое «{DICT_COLUMN}»', dictionary_note(morph)],
        ],
        columns=['Показатель', 'Значение'],
    )
    skipped_df = pd.DataFrame(
        [
            (word, count, 'мусор распознавания' if word in junk else 'редкое, но знакомое ученику')
            for word, count in sorted(stats_skipped.items(), key=lambda x: (x[0] not in junk, -x[1], x[0]))
        ] + [
            (word, count, f'образовано от знакомого: {base}')
            for word, (count, base) in sorted(stats_derived.items())
        ],
        columns=['Лемма', 'Вхождений в учебнике', 'Причина'],
    )

    with pd.ExcelWriter(str(out_path), engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Сводка', index=False)
        new_df.to_excel(writer, sheet_name='Новые слова', index=False)
        terms_df.to_excel(writer, sheet_name='Термины', index=False)
        skipped_df.to_excel(writer, sheet_name='Не выделено', index=False)
    print(f"[Отчёт] Сохранён: {out_path}")


# =============================================================================
# Пайплайн
# =============================================================================

def highlight_textbook(
    in_path: str | Path,
    grade: int,
    out_path: str | Path | None = None,
    base: str = 'both',
    dict_path: str | Path = DICT_PATH,
    include_propn: bool = False,
    use_cache: bool = True,
    add_legend: bool = True,
    min_freq: int = MIN_KNOWN_FREQ,
    not_new_path: str | Path = NOT_NEW_WORDS_PATH,
    familiar_path: str | Path = FAMILIAR_WORDS_PATH,
) -> Path:
    """
    Размечает учебник двумя цветами и пишет отчёт.

    Args:
        in_path: исходный docx
        grade: класс учебника (база сравнения — все классы младше)
        out_path: куда сохранить размеченный docx
        base: источник знакомых слов — 'rulex' | 'textbooks' | 'both'
        dict_path: словарь терминов
        include_propn: считать имена собственные новой лексикой
        use_cache: использовать кэш лемм предыдущих классов
        add_legend: записать пояснение к выделению отдельным файлом рядом с учебником
        min_freq: сколько упоминаний в предыдущих классах делает слово знакомым
        not_new_path: файл лемм, которые не выделяются как новые (мусор распознавания)
        familiar_path: файл редких, но знакомых ученику слов, которые не выделяются

    Returns:
        Путь к сохранённому файлу
    """
    in_path = Path(in_path)
    out_path = Path(out_path) if out_path else (
        HIGHLIGHT_DIR / f"Порог {min_freq}" / f"{_output_stem(in_path)}_подсветка_порог{min_freq}.docx"
    )

    print("=" * 70)
    print(f"Учебник:         {in_path.name}")
    print(f"Класс:           {grade}")
    print(f"База сравнения:  {base}")
    print(f"Результат:       {out_path}")
    print("=" * 70)

    preprocessor = TextPreprocessorWithoutStopwords(
        stopwords_path=STOPWORDS_PATH,
        replacements_path=REPLACEMENTS_PATH,
    )

    prev, sources = known_lemmas(
        grade, base, preprocessor, use_cache=use_cache, min_freq=min_freq
    )
    if min_freq > 1:
        base_descr = f"в предыдущих классах ({sources}) они встретились меньше {min_freq} раз"
    else:
        base_descr = f"в предыдущих классах ({sources}) их нет"
    junk = load_not_new_words(not_new_path)
    familiar = load_not_new_words(familiar_path)
    not_new = junk | familiar
    import term_dictionaries
    patterns_by_len, verb_shadowed, term_sources, term_lists, term_of = \
        term_dictionaries.load_all_patterns(grade, preprocessor)
    patterns_total = sum(len(v) for v in patterns_by_len.values())

    doc = Document(str(in_path))
    print(f"[Цвет издателя] снято с фрагментов: {strip_publisher_colors(doc)}")
    a = analyze_document(
        doc, preprocessor, prev, patterns_by_len, verb_shadowed, not_new,
        include_propn=include_propn, term_sources=term_sources,
    )
    t0 = a.t0
    paragraphs, stream = a.paragraphs, a.stream
    colored_by_para = a.colored_by_para
    stats_new, stats_terms = a.stats_new, a.stats_terms
    stats_skipped, stats_derived = a.stats_skipped, a.stats_derived

    import textbook_report
    a.update(
        in_path=in_path, grade=grade, min_freq=min_freq, known=prev, sources=sources,
        junk=junk, patterns_by_len=patterns_by_len, patterns_total=patterns_total,
        preprocessor=preprocessor,
        verb_shadowed=verb_shadowed, term_sources=term_sources, term_lists=term_lists, term_of=term_of,
        new_terms={
            term for term in stats_terms
            if any(not is_familiar(l, prev, preprocessor) for l in term.split())
        },
    )
    # Пометки у терминов: сколько раз термин встречался в прошлых классах
    notes = textbook_report.term_annotations(a, every=NOTE_MODE == 'comments')

    painted = 0
    for pi, spans in colored_by_para.items():
        painted += apply_colors_to_paragraph(paragraphs[pi], sorted(spans, key=lambda x: x[0]))
    noted = 0
    if NOTE_MODE == 'comments':
        comment_style(doc)
    for pi, items in notes["by_para"].items():
        noted += add_annotations(paragraphs[pi], items)
    print(f"[Пометки] Чисел упоминаний у терминов: {noted}")

    if add_legend:
        legend_path = out_path.with_name(f"Пояснение к выделению_{grade} класс_порог{min_freq}.docx")
        write_legend(legend_path, grade, base_descr, min_freq, notes["bio_span"], notes["all_span"])
        print(f"[Пояснение] {legend_path}")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))

    print("-" * 70)
    print(f"[Готово] Перекрашено фрагментов: {painted}")
    print(
        f"[Готово] Новых слов: {len(stats_new)} уникальных / "
        f"{sum(stats_new.values())} вхождений"
    )
    print(
        f"[Готово] Терминов: {len(stats_terms)} уникальных / "
        f"{sum(stats_terms.values())} вхождений"
    )
    for label, in_group in (
        ('мусор распознавания', lambda w: w in junk),
        ('редкие, но знакомые', lambda w: w not in junk),
    ):
        part = {w: n for w, n in stats_skipped.items() if in_group(w)}
        print(
            f"[Готово] Не выделено ({label}): {len(part)} лемм / "
            f"{sum(part.values())} вхождений"
        )
    print(
        f"[Готово] Не выделено (образованы от знакомых): {len(stats_derived)} лемм / "
        f"{sum(n for n, _ in stats_derived.values())} вхождений"
    )
    print(f"[Готово] Файл: {out_path}  ({round(time.time() - t0)} с)")

    # Подробный отчёт: страницы, расстояния, определения, параграфы и т. д.
    textbook_report.write_excel_report(
        a,
        HIGHLIGHT_DIR / f"Порог {min_freq}" / f"{_output_stem(in_path)}_отчёт_порог{min_freq}.xlsx",
        corpus=notes["corpus"],
    )
    return out_path


class Analysis(dict):
    """Результаты разбора учебника; поля доступны и как атрибуты."""
    __getattr__ = dict.__getitem__
    __setattr__ = dict.__setitem__


def analyze_document(
    doc,
    preprocessor,
    prev: set[str],
    patterns_by_len: dict[int, list[str]],
    verb_shadowed: dict[str, set[str]],
    not_new: set[str],
    include_propn: bool = False,
    pretokenized: Analysis | None = None,
    term_sources: dict | None = None,
) -> Analysis:
    """
    Находит в документе термины и новые слова, не меняя его.

    Сначала весь поток токенизируется: так собираются леммы, которые
    встречаются как нарицательные (common_lemmas), — по ним слово
    в начале предложения, помеченное как имя собственное, не теряется.

    Returns:
        Analysis с полями paragraphs, texts, stream, blocks
        [(начало, конец, токены)], occurrences [(вид, ключ, начало, конец)]
        в координатах потока, colored_by_para и счётчиками stats_*.
    """
    t0 = time.time()
    if pretokenized is not None:
        # Тот же документ с другим порогом: токены от порога не зависят
        paragraphs, texts = pretokenized.paragraphs, pretokenized.texts
        stream, blocks = pretokenized.stream, pretokenized.blocks
        common_lemmas = pretokenized.common_lemmas
        return _spans_pass(t0, paragraphs, texts, stream, blocks, common_lemmas, preprocessor,
                           prev, patterns_by_len, verb_shadowed, not_new, include_propn, term_sources)
    paragraphs = list(iter_paragraphs(doc))
    texts = [paragraph_run_text(p) for p in paragraphs]
    drop_caps = TextStream.find_drop_caps(texts, preprocessor.morph)
    stream = TextStream(texts, drop_caps, morph=preprocessor.morph)
    print(
        f"[Разметка] Абзацев: {len(paragraphs)}, символов: {len(stream.text)}, "
        f"сшито переносов: {stream.joins} (с дефисом: {stream.kept_hyphens}), "
        f"буквиц: {len(drop_caps)}"
    )

    blocks = []
    for b_start, b_end in stream.blocks():
        blocks.append((b_start, b_end, tokens_with_punct(stream.text[b_start:b_end], preprocessor)))
        print(f"[Токены] {b_end}/{len(stream.text)} символов ({round(time.time() - t0)} с)")
    common_lemmas = {
        t['lemma'] for _, _, tokens in blocks for t in tokens
        if t.get('pos') in NEW_WORD_POS and t.get('lemma')
    }
    return _spans_pass(t0, paragraphs, texts, stream, blocks, common_lemmas, preprocessor,
                       prev, patterns_by_len, verb_shadowed, not_new, include_propn, term_sources)


def _spans_pass(t0, paragraphs, texts, stream, blocks, common_lemmas, preprocessor,
                prev, patterns_by_len, verb_shadowed, not_new, include_propn, term_sources=None) -> Analysis:
    """Поиск терминов и новых слов по готовым токенам (см. analyze_document)."""
    a = Analysis(
        t0=t0, paragraphs=paragraphs, texts=texts, stream=stream, blocks=blocks,
        common_lemmas=common_lemmas,
        occurrences=[], colored_by_para=defaultdict(list),
        stats_new=Counter(), stats_terms=Counter(), stats_skipped=Counter(), stats_derived={},
    )
    # Интервалы считаются по потоку блоками, затем раскладываются по абзацам
    for b_start, b_end, tokens in blocks:
        occ: list = []
        colored = compute_spans(
            tokens,
            patterns_by_len,
            prev,
            preprocessor,
            a.stats_terms,
            a.stats_new,
            include_propn=include_propn,
            verb_shadowed=verb_shadowed,
            not_new=not_new,
            stats_skipped=a.stats_skipped,
            stats_derived=a.stats_derived,
            raw_text=stream.text[b_start:b_end],
            occurrences=occ,
            common_lemmas=common_lemmas,
            term_sources=term_sources,
        )
        a.occurrences.extend((kind, key, b_start + s, b_start + e) for kind, key, s, e in occ)
        for s, e, color in colored:
            for pi, ps, pe in stream.to_paragraphs(b_start + s, b_start + e):
                a.colored_by_para[pi].append((ps, pe, color))
    print(f"[Разметка] готово ({round(time.time() - t0)} с)")
    return a


def _output_stem(in_path: Path) -> str:
    """Имя исходного файла без служебной приписки «toc_» для имён результатов."""
    stem = in_path.stem
    return stem[len('toc_'):] if stem.startswith('toc_') else stem


def load_not_new_words(path: str | Path) -> set[str]:
    """
    Читает список лемм, которые не надо выделять как новые слова
    (not_new_words.txt — мусор распознавания, familiar_words.txt —
    редкие, но знакомые ученику слова).

    Формат: одна лемма на строку (как в столбце «Новое слово (лемма)»
    отчёта), всё после # — комментарий.
    """
    path = Path(path)
    if not path.exists():
        return set()
    words = set()
    for line in path.read_text(encoding='utf-8').splitlines():
        word = line.split('#', 1)[0].strip().lower().replace('ё', 'е')
        if word:
            words.add(word)
    print(f"[Исключения] {path.name}: {len(words)} лемм")
    return words


def write_legend(out_path: str | Path, grade: int, base_descr: str, min_freq: int,
                 bio_span: str = '', all_span: str = '') -> Path:
    """
    Пояснение к выделению — отдельным файлом рядом с размеченным учебником
    (в самом учебнике его нет, чтобы не сдвигать страницы).
    """
    import term_dictionaries as td
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    doc.add_heading(f"Пояснение к выделению в учебнике биологии, {grade} класс", level=1)

    def colored(par, text, color=None, ucolor=None, double=False, bold=True):
        run = par.add_run(text)
        run.bold = bold
        _apply_style(run, (color, ucolor, double))
        return run

    p = doc.add_paragraph()
    p.add_run("Цвет терминов: ").bold = True
    for i, src in enumerate(td.SOURCES):
        label = {"frp": "термины ФРП (рабочей программы)", "index": "указатель терминов учебника",
                 "dict": "общий словарь терминов"}[src]
        colored(p, td.COLOR_NAMES[src], td.FONT_COLOR[src])
        p.add_run(f" — {label}" + ("; " if i < len(td.SOURCES) - 1 else ". "))
    p.add_run("Термин из нескольких словарей: цвет слова — по старшему словарю (ФРП → указатель → "
              "словарь терминов), подчёркивание — цветом второго словаря (")
    colored(p, "зелёное", None, td.UNDERLINE_COLOR["index"], bold=False)
    p.add_run(" — указатель, ")
    colored(p, "жёлтое", None, td.UNDERLINE_COLOR["dict"], bold=False)
    p.add_run(" — словарь терминов), ")
    colored(p, "двойное", None, td.UNDERLINE_COLOR["index"], True, bold=False)
    p.add_run(" — термин во всех трёх словарях.")

    p = doc.add_paragraph()
    colored(p, "Бирюзовый", COLOR_NEW)
    p.add_run(f" — слова, новые для ученика {grade} класса: {base_descr}. Новизна определяется по "
              "лемме (начальной форме слова): все формы слова считаются вместе, а «раз» — это "
              "вхождения (токены) любой формы.")

    if bio_span and NOTE_MODE == 'comments':
        p = doc.add_paragraph()
        p.add_run("Числа у термина ").bold = True
        p.add_run("— сколько раз термин встречался в прошлых классах (вхождений, токенов), по классам: "
                  "«5:12» значит «в 5 классе 12 раз». ")
        colored(p, "био", NOTE_BIO_COLOR)
        p.add_run(f" — учебники биологии {bio_span} классов, ")
        colored(p, "все", NOTE_ALL_COLOR)
        p.add_run(f" — все школьные учебники {all_span} классов. Классы, где термина не было, не показаны; "
                  f"«0» — не встречался ни разу; «т» — тысяч.")
        p = doc.add_paragraph()
        p.add_run("Где смотреть числа. ").bold = True
        p.add_run("В учебнике Word числа стоят в примечаниях: у каждого термина при первом его появлении на "
                  "каждой странице (повторы на той же странице без примечаний); соседние термины собраны в одну "
                  "выноску, по строке на термин (вкладка «Рецензирование» → «Показать примечания» или область примечаний). "
                  "Примечания не занимают места в тексте, поэтому страницы совпадают с исходным учебником. "
                  "Для чтения и печати есть отдельный файл «…числа_у_терминов….pdf»: слева страница учебника, "
                  "справа — все термины этой страницы с числами.")
    elif bio_span:
        p = doc.add_paragraph()
        p.add_run("Числа у термина ").bold = True
        p.add_run(f"(при первом упоминании и там, где он снова появляется после перерыва в 20 и больше "
                  f"страниц) — сколько раз термин встречался в прошлых классах (вхождений, токенов), по "
                  f"классам: «5:12» значит «в 5 классе 12 раз». ")
        colored(p, "Верхняя строка", NOTE_BIO_COLOR)
        p.add_run(f" (над термином) — учебники биологии {bio_span} классов, ")
        colored(p, "нижняя", NOTE_ALL_COLOR)
        p.add_run(f" (под термином) — все школьные учебники {all_span} классов. Классы, где термина не "
                  f"было, не показаны; «0» — не встречался ни разу; «т» — тысяч. Числа стоят в "
                  f"плавающих надписях и не сдвигают текст: страницы совпадают с исходным учебником.")

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Подсветка терминов и новой лексики в учебнике docx",
    )
    parser.add_argument('--in', dest='in_path', required=True, help="исходный docx")
    parser.add_argument('--grade', type=int, required=True, help="класс учебника")
    parser.add_argument('--out', dest='out_path', default=None, help="выходной docx")
    parser.add_argument(
        '--base',
        choices=('both', 'rulex', 'textbooks'),
        default='both',
        help="база знакомых слов: списки RuLex 2..N-1, учебники биологии или оба",
    )
    parser.add_argument(
        '--include-propn',
        action='store_true',
        help="считать имена собственные новой лексикой",
    )
    parser.add_argument(
        '--min-freq',
        type=int,
        default=MIN_KNOWN_FREQ,
        help=f"сколько упоминаний в предыдущих классах делает слово знакомым "
             f"(по умолчанию {MIN_KNOWN_FREQ})",
    )
    parser.add_argument('--no-cache', action='store_true', help="не использовать кэш лемм")
    parser.add_argument('--no-legend', action='store_true', help="не записывать пояснение к выделению")
    args = parser.parse_args()

    highlight_textbook(
        in_path=args.in_path,
        grade=args.grade,
        out_path=args.out_path,
        base=args.base,
        include_propn=args.include_propn,
        use_cache=not args.no_cache,
        add_legend=not args.no_legend,
        min_freq=args.min_freq,
    )


if __name__ == "__main__":
    main()
