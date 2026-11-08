"""
Подробный отчёт по размеченному учебнику.

Excel: новые слова и термины со страницами, расстояниями между вхождениями,
определениями в учебнике и в других учебниках, пометкой «биологический ли
термин», упоминаниями в прошлых классах; плотность терминов по параграфам.

PDF-черновик для издателя (write_pdf_report): слева страница учебника
с подсветкой, справа рекомендации по этой странице.

Страницы берутся из PDF издательства: каждый абзац docx находится в
текстовом слое PDF. Параграфы (§) — по оглавлению PDF; конец параграфа
с рубрики «Тренируемся» (вопросы, задания, «Это интересно», «Шаги
к успеху» и т. д.) в плотность не входит.

Запуск (только отчёты, docx не пересобирается):
    python textbook_report.py --grade 10 --min-freq 3 [--pdf]
"""
from __future__ import annotations

import argparse
import json
import logging
import pickle
import re
import statistics
from array import array
from bisect import bisect_left
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import config
import highlight_new_and_terms as h
from src.analyzer import build_indexed_words

logging.getLogger("pdfminer").setLevel(logging.ERROR)

CACHE_DIR = h.CACHE_DIR
BIO_SUBJECTS = {"Биология", "Биология (доп.)"}
# Предметы с маленьким корпусом (музыка, ИЗО…) дают случайные частоты
MIN_SUBJECT_WORDS = 100_000
# Биологический термин: на биологию приходится не меньше четверти частоты
# слова (частоты на миллион слов по предметам складываются; если бы слово
# одинаково часто встречалось во всех ~16 предметах, доля была бы ~6 %)
BIO_SHARE = 0.25
# Сколько лет без упоминаний — «давно не встречалось»
LONG_GAP_YEARS = 4
# Вхождения одного слова дальше этого числа страниц — «далеко друг от друга»
FAR_PAGES = 20

WORD_RE = re.compile(r"[а-яё]+(?:-[а-яё]+)*")


def compact(text: str) -> str:
    """Только буквы и цифры в нижнем регистре: так сравниваются docx и PDF."""
    return re.sub(r"[^0-9a-zа-я]", "", text.lower().replace("ё", "е"))


def raz(n: int) -> str:
    """1 раз, 2 раза, 5 раз."""
    return "раза" if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14) else "раз"


def pages_str(pages) -> str:
    """[5, 6, 7, 9] -> «5–7, 9»."""
    pages = sorted({p for p in pages if p})
    out, i = [], 0
    while i < len(pages):
        j = i
        while j + 1 < len(pages) and pages[j + 1] == pages[j] + 1:
            j += 1
        out.append(str(pages[i]) if i == j else f"{pages[i]}–{pages[j]}")
        i = j + 1
    return ", ".join(out)


# =============================================================================
# PDF: текст страниц и оглавление
# =============================================================================

def load_pdf(pdf_path: Path, use_cache: bool = True) -> dict:
    """
    Текст и слова (с координатами) каждой страницы PDF. Страница PDF с
    индексом i — страница учебника i + 1 (проверено по колонтитулам).
    """
    pdf_path = Path(pdf_path)
    cache = CACHE_DIR / f"pdf_{pdf_path.stem}.json"
    if use_cache and cache.exists() and cache.stat().st_mtime >= pdf_path.stat().st_mtime:
        return json.loads(cache.read_text(encoding="utf-8"))
    import pdfplumber
    pages = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            pages.append({
                "text": page.extract_text() or "",
                "w": float(page.width),
                "h": float(page.height),
                "words": [[w["text"], round(w["x0"], 1), round(w["top"], 1),
                           round(w["x1"], 1), round(w["bottom"], 1)] for w in words],
            })
    data = {"pages": pages}
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print(f"[PDF] {pdf_path.name}: {len(pages)} страниц")
    return data


TOC_ENTRY_RE = re.compile(r"^(.*?)[\s.…]+(\d{1,3})$")


def parse_toc(pdf: dict) -> tuple[list[dict], list[int]]:
    """
    Оглавление PDF: параграфы [{num, title, start}] и начальные страницы
    всех разделов (введение, главы, приложения, указатели) — границы,
    на которых параграф заканчивается.
    """
    pages = pdf["pages"]
    first = next(
        i for i, p in enumerate(pages)
        if "Оглавление" in p["text"] and p["text"].count("§") >= 5
    )
    lines = []
    for p in pages[first:first + 3]:
        if p is not pages[first] and "§" not in p["text"]:
            break
        lines += p["text"].splitlines()

    sections, starts, buf = [], [], ""
    for line in lines:
        line = line.strip()
        if not line or line == "Оглавление" or (not buf and re.fullmatch(r"\d{1,3}", line)):
            continue
        buf = (buf[:-1] + line if buf.endswith("-") else f"{buf} {line}") if buf else line
        m = TOC_ENTRY_RE.match(buf)
        # Номер страницы — после отточия или просто в конце строки, если он
        # больше предыдущего («…Экологическая ниша 150»)
        if not m or not (re.search(r"(\.\s?){3,}|…", buf) or (starts and int(m.group(2)) > starts[-1])):
            continue
        title = re.sub(r"[\s.…]+$", "", m.group(1)).strip()
        start = int(m.group(2))
        starts.append(start)
        sm = re.match(r"§\s*(\d+)\.\s*(.*)", title)
        if sm:
            sections.append({"num": int(sm.group(1)), "title": sm.group(2).strip(), "start": start})
        buf = ""
    starts = sorted(set(starts))
    for s in sections:
        later = [x for x in starts if x > s["start"]]
        s["end"] = (later[0] - 1) if later else len(pages)
    return sections, starts


class PageMap:
    """
    Страница учебника для каждого абзаца docx и для любой позиции потока.

    Абзац ищется в тексте PDF по трём отрывкам (начало, середина, конец)
    сначала рядом с предыдущей найденной страницей, потом шире. Позиция
    внутри абзаца уточняется по 14 символам текста, начиная с неё: абзац
    мог начаться на одной странице и закончиться на следующей.
    """

    def __init__(self, a, pdf: dict):
        self.pages = [compact(p["text"]) for p in pdf["pages"]]
        self.stream = a.stream
        n = len(self.pages)
        para_page: list[int | None] = [None] * len(a.texts)
        last = 1
        for i, t in enumerate(a.texts):
            c = compact(t)
            if len(c) < 12:
                continue
            page = self._find(c, last)
            if page:
                para_page[i] = last = page
        cur = None
        for i, p in enumerate(para_page):
            if p is None:
                para_page[i] = cur
            else:
                cur = p
        self.para_page = para_page
        self.n = n

    def _find(self, c: str, last: int) -> int | None:
        mid = len(c) // 2
        snippets = [c[:16], c[mid:mid + 16], c[-16:]]
        near = [last, last + 1, last + 2, last + 3, last - 1]
        wide = list(range(last - 3, last + 13))
        for cand, min_len in ((near, 12), (wide, 12), (range(1, len(self.pages) + 1), 30)):
            if len(c) < min_len:
                continue
            for p in cand:
                if 1 <= p <= len(self.pages) and any(s in self.pages[p - 1] for s in snippets):
                    return p
        return None

    def para_of(self, pos: int) -> int | None:
        origin = self.stream.origin
        for k in range(pos, min(pos + 200, len(origin))):
            if origin[k] is not None:
                return origin[k][0]
        return None

    def page_of(self, pos: int) -> int | None:
        pi = self.para_of(pos)
        if pi is None:
            return None
        p = self.para_page[pi]
        if p is None:
            return None
        c = compact(self.stream.text[pos:pos + 80])[:14]
        if len(c) >= 6:
            for q in (p, p + 1, p - 1, p + 2):
                if 1 <= q <= self.n and c in self.pages[q - 1]:
                    return q
        return p


# =============================================================================
# Слова потока, параграфы, плотность
# =============================================================================

def word_starts(a) -> list[int]:
    """Начала всех слов потока (по порядку): номер слова = индекс в списке."""
    starts = []
    for b_start, _, tokens in a.blocks:
        starts += [b_start + w["start"] for w in build_indexed_words(tokens)]
    return starts


def find_sections(a, pm: PageMap, pdf: dict) -> list[dict]:
    """
    Параграфы с границами основного текста.

    Параграф — страницы от его начала по оглавлению до начала следующего
    раздела. Основной текст — абзацы этих страниц до рубрики «Тренируемся»;
    сама рубрика и всё после неё (вопросы, «Совершенствуемся», «Обсуждаем»,
    «Это интересно», «Шаги к успеху»…) в плотность не входят.
    """
    sections, _ = parse_toc(pdf)
    trainer = [i for i, t in enumerate(a.texts) if t.strip().startswith("Тренируемся") and len(t.strip()) < 20]
    for s in sections:
        paras = [i for i, p in enumerate(pm.para_page) if p and s["start"] <= p <= s["end"]]
        s["first_para"] = paras[0] if paras else None
        s["last_para"] = paras[-1] if paras else None
        tr = [i for i in trainer if pm.para_page[i] and s["start"] <= pm.para_page[i] <= s["end"]]
        s["trainer_para"] = tr[0] if tr else None
        s["trainer_page"] = pm.para_page[tr[0]] if tr else None
    return sections


def section_of_para(sections: list[dict], pm: PageMap, pi: int | None) -> tuple[int | None, bool]:
    """(номер параграфа, входит ли абзац в основной текст параграфа)."""
    if pi is None:
        return None, False
    page = pm.para_page[pi]
    for s in sections:
        if page and s["start"] <= page <= s["end"]:
            main = s["trainer_para"] is None or pi < s["trainer_para"]
            return s["num"], main
    return None, False


# =============================================================================
# Склейка вариантов леммы: «атф-синтетаз» и «атф-синтетаза»
# =============================================================================

def merge_variants(surfaces: dict[str, Counter], counts: Counter, morph, form_count=None) -> dict[str, str]:
    """
    Natasha иногда даёт одному слову разные леммы в разных падежах
    («АТФ-синтетазы» -> «атф-синтетаз», «АТФ-синтетазой» -> «атф-синтетаза»).
    Леммы склеиваются, если у них общая начальная форма pymorphy3 хотя бы
    одной словоформы или общая основа без падежного окончания (6+ букв).

    Название группы — начальная форма, которая чаще встречается как
    словоформа в учебниках всех предметов (form_count); если её там нет —
    та, из которой pymorphy3 выводит больше всего словоформ текста.

    Returns:
        {лемма: название группы}
    """
    def norms(surface: str) -> set[str]:
        low = surface.lower().replace("ё", "е")
        head, sep, last = low.rpartition("-")
        # Догадки pymorphy3 «это имя / организация» для научных слов ложные
        return {head + sep + p.normal_form.replace("ё", "е") for p in morph.parse(last)
                if not ({"Name", "Surn", "Patr", "Orgn", "Geox"} & p.tag.grammemes)}

    parent = {l: l for l in surfaces}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    owner: dict[str, str] = {}
    norms_of: dict[str, dict[str, set[str]]] = {}
    for lemma, forms in surfaces.items():
        norms_of[lemma] = {f: norms(f) for f in forms}
        keys = {lemma} | set().union(*norms_of[lemma].values())
        keys |= {k2 for k in keys | set(forms) for k2 in stem_keys(k)}
        for k in keys:
            if k in owner:
                parent[find(lemma)] = find(owner[k])
            else:
                owner[k] = lemma
    groups = defaultdict(list)
    for l in surfaces:
        groups[find(l)].append(l)
    canon = {}
    for members in groups.values():
        support = Counter({m: 0 for m in members})
        for m in members:
            for f, ns in norms_of[m].items():
                for cand in ns:
                    support[cand] += surfaces[m][f]
        fc = form_count or (lambda w: 0)
        best = max(support, key=lambda c: (fc(c), support[c], c in members, -len(c)))
        for m in members:
            canon[m] = best
    return canon


# Окончания, которыми различаются варианты одной леммы у Natasha
# («аутофагосом» / «аутофагосома», «пируват» / «пирувата», «бислой» / «бислый»)
INFLECTIONS = sorted({
    "", "а", "я", "ы", "и", "у", "ю", "е", "о", "ом", "ем", "ой", "ей", "ый", "ий", "ое", "ее",
    "ая", "яя", "ые", "ие", "ов", "ев", "ам", "ям", "ами", "ями", "ах", "ях", "ым", "им", "ых", "их",
}, key=len, reverse=True)


def stem_keys(word: str) -> set[str]:
    """Возможные основы без падежного окончания (6+ букв) — для склейки вариантов."""
    return {"~" + word[: len(word) - len(e)] for e in INFLECTIONS
            if word.endswith(e) and len(word) - len(e) >= 6}


def display_lemma(lemma: str, morph) -> str:
    """
    Начальная форма для отчёта: Natasha иногда оставляет косвенный падеж
    («пирувата», «олигосахарида») — тогда берётся начальная форма pymorphy3.
    """
    head, sep, last = lemma.rpartition("-")
    p = morph.parse(last)[0]
    nf = p.normal_form.replace("ё", "е")
    if nf == last or p.tag.POS not in ("NOUN", "ADJF"):
        return lemma
    if "nomn" not in p.tag or (p.tag.POS == "ADJF" and "masc" not in p.tag) or (p.tag.POS == "NOUN" and "plur" in p.tag):
        return head + sep + nf
    return lemma


# =============================================================================
# Определения
# =============================================================================

QUOTE = r"[»\"”]?"
LEAD = rf"^\s*{QUOTE}\s*(?:\([^)]{{0,160}}\)\s*)?(?:,\s*или\s+[^,—–]{{1,60}},\s*)?"
NAMING_BEFORE = re.compile(
    r"(называ\w*|назва\w*|счита\w*|получил\w*\s+(?:\w+\s+)?назван\w*|носит\w*\s+назван\w*|именуют\w*|"
    r"обознача\w*\s+(?:\w+\s+)?(?:термином|понятием)|термином|понятием)\s+(?:[\w-]+\s+){0,2}[«\"“]?\s*$"
)
# После тире — родовое понятие определения, перед ним может быть до двух
# прилагательных: «геномика — наука, изучающая…», «ретротранспозоны —
# подвижные участки генома», «онихофоры — интересная группа животных»
GENUS_AFTER_DASH = re.compile(
    r"\s*(?:[а-яё]+(?:ые|ый|ий|ой|ая|яя|ое|ее|ие|ых|их|ую|юю|ого|его|ому|ему|ыми|ими)\s+){0,2}"
    r"(?:это\b|наук|процесс|способн|совокупн|веществ|соединени|фермент|бел[оо]?к|органоид|структур|"
    r"явлени|метод|част[ьи]|групп|форм[аы]|особ[аоыи]|специальн|тип\b|вид\b|комплекс|молекул|клет[ко]|"
    r"участ|раздел|отрасл|направлени|свойств|состояни|образовани|пузыр|углевод|липид|кислот|орган\b|ткан|"
    r"бактери|\w*полимер|прибор|предк|предок|потомк|родственник|рецептор|учени|животн|растени|организм|"
    r"гриб|черв|рыб|рептили|ящер|динозавр|гормон|пигмент|порок|заболевани|болезн|ген\b|гены|слой|оболочк|"
    r"мембран|нит[ьи]|сахар|спирт|элемент|частиц|фаз|период|стади|этап|реакци|механизм|систем|отдел|"
    r"класс|отряд|семейств|род\b|царств|сообществ|популяци|признак|продукт|препарат|фактор|вирус|"
    r"водоросл|насеком|птиц|млекопитающ|амфиби|земноводн|пресмыкающ|моллюск|растворител|катализатор)"
)
# После тире — именная группа и причастный или придаточный оборот:
# «остеокластов — клеток, разрушающих кость», «ЯМР-спектрографа — прибора, измеряющего…»
APPOSITION_AFTER_DASH = re.compile(
    r"\s*[^,.;:!?()—–]{3,70},\s*(?:котор\w+|[а-яё]+(?:ющ|ащ|ящ|вш|ем|им|нн|т)(?:ий|ая|ее|ие|его|ей|их|ую|ими|ым)(?:ся)?\b)"
)
# Начало пункта перечня или скобки: «…; мезолецитальные — содержат…», «(родопсин — рецептор…)»
CLAUSE_START = re.compile(r"(?:^|[;:(]|\b\d{1,2}[).])\s*[«\"“]?\s*$")
# «Другие амиды — твёрдые вещества»: высказывание о части класса, а не определение
NOT_DEF_BEFORE = re.compile(
    r"(?:\bне|\bа|други[ехм]|некотор\w+|многи[ехм]|\bвсе|\bэти|\bтакие|\bоба|\bобе|\bодни|\bчасть|\bиз)\s+[«\"“]?\s*$")
# «а денисовские — у обитателей…»: тире заменяет пропущенный глагол
PREP_AFTER_DASH = re.compile(r"\s*(?:у|в|во|на|из|от|с|со|к|по|при|за|до|для|над|под)\s")
PRONOUN_AFTER_DASH = re.compile(r"\s*(?:он[аио]?|они|его|её|их|ему|ей|им)\b")
# Слово — само слово-источник: «(от греч. атомос — неделимый)»
ETYMON_BEFORE = re.compile(r"\(\s*от\s+(?:др\.\s*-?\s*)?[а-яё]+\.?\s+(?:[а-яё]+\s+)?$")
INSTRUMENTAL = re.compile(r"(?:ой|ей|ом|ем|ами|ями|ою|ею)$")
# Родовое слово прямо перед X: «ферментом АТФ-синтетазой, которая…»,
# «гетерополимера муреина, в состав которого…»
GENUS_BEFORE = re.compile(
    r"(?:фермент|бел[оо]?к|белк|веществ|соединени|органоид|структур|клет[ко]|\w*полимер|молекул|"
    r"пигмент|гормон|рецептор|бактери|организм|животн|растени|гриб|участ|комплекс|процесс|"
    r"явлени|метод|прибор|кислот|углевод|сахар|липид)\w*\s+$")
RELATIVE_AFTER = re.compile(r"^\s*[»\"”]?\s*(?:\([^)]{0,80}\)\s*)?,\s*(?:(?:в|при|из|с|на|у|для)\s+(?:\w+\s+)?)?котор")
PARTICIPLE_AFTER = re.compile(
    r"^\s*,\s*(?:котор\w+|[а-яё]+(?:ющ|ащ|ящ|вш|ем|им|нн|т)(?:ий|ая|ее|ие|его|ей|их|ую|ими|ым|ые)(?:ся)?\b)")
# Классификация: «разделяют на гомополимеры, которые…», «Это фибриллярные белки, которые…»
CLASSIFY_BEFORE = re.compile(
    r"(?:(?:под)?раздел\w+|дел[яи]т\w*|различа\w+|выделя\w+)\s+(?:\w+\s+){0,3}на\s+(?:[\w-]+\s+){0,2}$"
    r"|(?:^|[.;:]\s*|\bи\s+)это\s+$|,\s*и\s+$")
GLOSS_AFTER = re.compile(r"^\s*\((?!\s*(?:рис|см\.|табл|с\.|\d))\s*[а-яё]")
SO_CALLED = re.compile(r"так\s+называем\w*\s+(?:\w+\s+){0,2}[«\"“]?\s*$")
DEF_AFTER = re.compile(LEAD + r",?\s*(называ\w+|представля\w+\s+собой|означа\w+|определя\w+\s+как|"
                       r"понимают\s+как|включа\w+\s+в\s+себя)\b")
DASH_AFTER = re.compile(LEAD + r"[—–]\s*")
DASH_THIS = re.compile(LEAD + r"[—–]\s*это\b")
EXPLAIN_AFTER = re.compile(
    rf"^\s*{QUOTE}\s*(?:,\s*|\(\s*)(?:т\.\s*е\.|то\s+есть|или\s)|"
    rf"^\s*{QUOTE}\s*\(\s*от\s+(?:греч|лат|англ|фр|нем|итал|араб|исп|голл|санскр)"
)


def classify_definition(sentence: str, ws: int, we: int) -> str | None:
    """
    Как предложение вводит слово, стоящее в [ws, we):
    'def' — определение: «X — это…»; «X — …» в начале предложения, пункта
    перечня или скобки; «X — <родовое понятие>…» и «X — <понятие>, который /
    причастие…» в любом месте; «…называют X»; «X называется / представляет
    собой / включает в себя…»; «под X понимают…»; «к X относят…»;
    'explain' — пояснение: «X, т. е. …», «X (т. е. …)», «X, или Y», «(или X)»,
    «X (расшифровка)», «описание (X)», «описание — X.», «X (от греч. …)»;
    None — просто упоминание.
    """
    if sentence.rstrip().endswith("?"):
        return None
    before, after = sentence[:ws].lower(), sentence[we:].lower()
    word = sentence[ws:we].lower()
    if ETYMON_BEFORE.search(before):
        return None
    words_before = len(re.findall(r"\w+", before))
    clause_start = words_before <= 6 or bool(CLAUSE_START.search(before))
    if DASH_THIS.match(after):
        return "def"
    m = DASH_AFTER.match(after)
    # «потомков куксониевых — риниофитов»: тире поясняет главное слово, а не X
    governed = re.search(r"(?:ых|их)$", word) and re.search(r"[а-яё]+(?:ов|ев|ей)\s+$", before)
    if m and not governed and not NOT_DEF_BEFORE.search(before) and not PRONOUN_AFTER_DASH.match(after, m.end()) \
            and not PREP_AFTER_DASH.match(after, m.end()):
        short = re.match(r"\s*[а-яё][а-яё\s-]{2,60}?(?:[.;)]|$)", after[m.end():])
        if (clause_start or GENUS_AFTER_DASH.match(after, m.end()) or APPOSITION_AFTER_DASH.match(after, m.end())
                or (short and len(short.group().split()) <= 5)):
            return "def"
    if NAMING_BEFORE.search(before) and not SO_CALLED.search(before):
        return "def"
    dm = DEF_AFTER.match(after)
    if dm and not (dm.group(1).startswith("называ") and INSTRUMENTAL.search(word)):
        return "def"
    if re.search(r"\bпод\s+[«\"“]?$", before) and re.match(rf"^\s*{QUOTE}[^.]{{0,60}}?понима", after):
        return "def"
    if re.search(r"\bк\s+$", before) and re.match(r"^\s*(?:[а-яё-]+\s+){0,2}относ", after):
        return "def"
    # «то, которое побольше, называют макронуклеусом, а поменьше — микронуклеусом»,
    # «из двух — дипептидом, из трёх — трипептидом»
    if re.search(r"\b(?:называ|назва)\w*[^;:]{0,80}[—–]\s*$", before):
        return "def"
    # «Это фибриллярные белки, которые имеют…», «разделяют на гомополимеры, которые состоят…»
    if (RELATIVE_AFTER.match(after) and GENUS_BEFORE.search(before)) or (
            CLASSIFY_BEFORE.search(before) and re.match(r"^\s*(?:[а-яё-]+\s*)?,\s*котор", after)):
        return "def"
    # «особых белков — гистонов, которые участвуют…», «сложные белки — гликопротеины, соединённые…»
    if re.search(r"[—–]\s*$", before) and PARTICIPLE_AFTER.match(after):
        return "def"
    # «Микротрофные вещества включают в себя…», «Ко второму типу относят
    # постзиготические механизмы, включающиеся…»
    if re.match(r"^\s*[а-яё-]+\s+(?:называ\w+|представля\w+\s+собой|включа\w+\s+в\s+себя)", after) \
            and re.search(r"(?:ый|ые|ая|ое|ие|ий|ых|ым)$", word):
        return "def"
    if re.search(r"\bотнос\w*\s+$", before) and (
            re.match(r"^\s*(?:[а-яё-]+\s*)?,\s*(?:т\.\s*е\.|котор)", after) or PARTICIPLE_AFTER.match(
                re.sub(r"^\s*[а-яё-]+", "", after, count=1))):
        return "def"
    # «тератогенами, т. е. веществами, вызывающими…» — «т. е.» + родовое понятие = определение
    te = re.match(r"^\s*(?:[а-яё-]+\s+)?[а-яё-]*\s*[,(]\s*(?:т\.\s*е\.|то\s+есть)", after)
    if te and GENUS_AFTER_DASH.match(after, te.end()):
        return "def"
    # «…, т. е. определение (ощущение)…» и т. п. — пояснение
    if EXPLAIN_AFTER.match(after) or re.search(r",\s*или\s+$", before):
        return "explain"
    # Придаточное сразу после слова: «эндосимбиоза, при котором бактерии поселились…»
    if RELATIVE_AFTER.match(after):
        return "explain"
    if re.match(r"^\s*(?:[а-яё-]+\s+)?[а-яё-]+\s*(?:,\s*|\(\s*)(?:т\.\s*е\.|то\s+есть|или\s)", after):
        return "explain"
    if GLOSS_AFTER.match(after) or (re.search(r"\(\s*(?:или\s+)?$", before) and re.match(r"^\s*\)", after)):
        return "explain"
    if re.search(r"[—–]\s*$", before) and re.match(r"^\s*(?:[.;,)—–]|$)", after):
        return "explain"
    return None


def sentence_bounds(text: str, s: int, e: int, para_breaks: bool) -> tuple[int, int]:
    """Границы предложения вокруг [s, e)."""
    lo, hi = max(0, s - 600), min(len(text), e + 600)
    stop = r"[.!?…]\s+(?=[А-ЯЁA-Z«(—–\d])"
    if para_breaks:
        stop += r"|\n"
    left = list(re.finditer(stop, text[lo:s]))
    start = lo + left[-1].end() if left else lo
    right = re.search(r"[.!?…](?=\s+[А-ЯЁA-Z«(—–\d]|\s*$)" + (r"|\n" if para_breaks else ""), text[e:hi])
    end = e + right.end() if right else hi
    return start, end


ANAPHORA = re.compile(r"\s*(?:Они|Он|Она|Оно|Такие|Такой|Такая|Такое|Такую|Таким|Эти|Это|Этот|Эта|Эту|Его|Их|Им|Ими)")


def with_context(text: str, ss: int, se: int, para_breaks: bool) -> str:
    """Предложение; если оно начинается с «Они / Такие / Это…» — вместе с предыдущим."""
    if ANAPHORA.match(text, ss) and ss > 1:
        pss, _ = sentence_bounds(text, ss - 2, ss - 1, para_breaks)
        return text[pss:se]
    return text[ss:se]


def clean_sentence(t: str, limit: int = 400) -> str:
    t = re.sub(r"\s+", " ", t).strip()
    return t if len(t) <= limit else t[:limit].rsplit(" ", 1)[0] + "…"


def emphasis_keys(paragraphs) -> set[str]:
    """
    Начертания, которыми в учебнике выделены слова: жирное (или жирное
    курсивное), если им набрано меньше 5 % текста. В 10 классе основной
    текст прямой, в 11 — после конвертации курсивный, поэтому доля
    считается по самому учебнику.
    """
    chars = Counter()
    for p in paragraphs:
        for r in p.runs:
            if r.text.strip():
                chars[("B" if r.bold else "") + ("I" if r.italic else "")] += len(r.text)
    total = sum(chars.values()) or 1
    return {k for k, n in chars.items() if "B" in k and n / total < 0.05}


def is_emphasized(a, pos: int, end: int, keys: set[str]) -> bool:
    for pi, ps, _ in a.stream.to_paragraphs(pos, end)[:1]:
        off = 0
        for r in a.paragraphs[pi].runs:
            if off <= ps < off + len(r.text):
                return (("B" if r.bold else "") + ("I" if r.italic else "")) in keys
            off += len(r.text)
    return False


FIG_STATUS = "нет определения, есть рисунок"
FIG_NUM_RE = re.compile(r"\b[Рр]ис\.\s*(\d{1,3})")


def figure_of(text: str, ss: int, se: int, s: int) -> str | None:
    """
    Рисунок при слове: слово стоит в подписи к рисунку («Рис. 34. Шапероны
    обеспечивают…», «Схема строения…: 5 — плазмида … Рис.») или предложение
    ссылается на рисунок («… (рис. 67)»). Возвращает «рис. 34» или «рисунок».
    """
    window = text[max(0, ss - 40):se + 12]
    m = FIG_NUM_RE.search(window)
    if m:
        return f"рис. {m.group(1)}"
    if re.search(r"\bРис\.|\bСхема\b|^\s*\d+\s*—", text[ss:se]):
        return "рисунок"
    return None


def find_definitions(a, occ_by_key: dict[str, list], pm: PageMap) -> dict[str, dict]:
    """
    Для каждого слова/термина: есть ли в учебнике определение, где и какое.

    Returns:
        {ключ: {status, text, page, first_page}}; status — «есть»,
        «только пояснение», «выделено жирным, определения нет», «нет»
    """
    text = a.stream.text
    lengths = [len(t) for t in a.texts if t.strip()]
    para_breaks = statistics.median(lengths) > 100 if lengths else True
    keys = emphasis_keys(a.paragraphs)
    result = {}
    for key, occs in occ_by_key.items():
        best = {"def": None, "explain": None}
        bold = None
        figures = []
        contexts, seen = [], set()
        # Обрывки в 1–3 буквы («оту» из «раб оту») — мусор, определение не ищется
        short_junk = len(re.sub(r"[^а-яёa-z]", "", key)) < 4
        for s, e in occs:
            ss, se = sentence_bounds(text, s, e, para_breaks)
            fig = figure_of(text, ss, se, s)
            if fig:
                label = f"{fig}, с. {pm.page_of(s)}"
                if label not in figures:
                    figures.append(label)
            kind = None if short_junk else classify_definition(text[ss:se], s - ss, e - ss)
            if kind and best[kind] is None:
                best[kind] = (clean_sentence(with_context(text, ss, se, para_breaks), 500), pm.page_of(s))
            if bold is None and is_emphasized(a, s, e, keys):
                bold = pm.page_of(s)
            # Все предложения со словом — чтобы проверить «нет определения» глазами;
            # повторы (подпись к рисунку дублируется при конвертации) убираются
            sent = re.sub(r"\s+", " ", text[ss:se]).strip()
            if "УКАЗАТЕЛЬ ТЕРМИНОВ" in sent or len(re.findall(r"[А-Яа-яЁё]+ \d{1,3}(?:,| [А-ЯЁ])", sent)) >= 5:
                continue
            if sent not in seen:
                seen.add(sent)
                lead = re.sub(r"\s+", " ", text[ss:s]).lstrip()
                contexts.append((pm.page_of(s), sent, len(lead), len(lead) + len(re.sub(r"\s+", " ", text[s:e]))))
        first_page = pm.page_of(occs[0][0]) if occs else None
        if best["def"]:
            status, (sent, page) = "есть", best["def"]
        elif best["explain"]:
            status, (sent, page) = "только пояснение", best["explain"]
        elif figures:
            # Отдельный случай: определения нет, но слово показано на рисунке
            status, sent, page = FIG_STATUS, "", None
        elif bold:
            status, sent, page = "выделено жирным, определения нет", "", bold
        else:
            status, sent, page = "нет", "", None
        result[key] = {"status": status, "text": sent, "page": page, "first_page": first_page,
                       "contexts": contexts, "figures": figures}
    return result


# =============================================================================
# Корпус учебников всех предметов
# =============================================================================

class Corpus:
    """
    Учебники всех предметов (2–11 классы), лемматизированные pymorphy3.
    Нужны для двух вещей: в каких предметах встречается термин
    (биологический ли он) и есть ли определение слова в других учебниках.
    """

    def __init__(self, morph, use_cache: bool = True):
        self.morph = morph
        self._lemma_cache: dict[str, str] = {}
        src = config.CORPUS_CHUNKS
        cache = CACHE_DIR / "corpus_lemmas.pkl"
        if use_cache and cache.exists() and cache.stat().st_mtime >= src.stat().st_mtime:
            with open(cache, "rb") as f:
                d = pickle.load(f)
        else:
            d = self._build(src)
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            with open(cache, "wb") as f:
                pickle.dump(d, f, protocol=pickle.HIGHEST_PROTOCOL)
        self.vocab: list[str] = d["vocab"]
        self.ids = {w: i for i, w in enumerate(self.vocab)}
        self.chunks: list[array] = d["chunks"]
        self.meta: list[tuple] = d["meta"]
        self.subject_words = Counter()
        for arr, (subj, *_rest) in zip(self.chunks, self.meta):
            self.subject_words[subj] += len(arr)
        self._texts = None

    def lemma(self, word: str) -> str:
        w = word.lower().replace("ё", "е")
        lem = self._lemma_cache.get(w)
        if lem is None:
            lem = self.morph.parse(w)[0].normal_form.replace("ё", "е")
            self._lemma_cache[w] = lem
        return lem

    def _build(self, src: Path) -> dict:
        print("[Корпус] лемматизация учебников всех предметов (один раз, потом из кэша)…")
        df = pd.read_parquet(src, columns=["text", "subject", "grade", "grade_max", "doc_id"])
        vocab, ids, chunks, meta = [], {}, [], []
        for text, subj, grade, gmax, doc in zip(df.text, df.subject, df.grade, df.grade_max, df.doc_id):
            arr = array("I")
            for w in WORD_RE.findall((text or "").lower()):
                lem = self.lemma(w)
                i = ids.get(lem)
                if i is None:
                    i = ids[lem] = len(vocab)
                    vocab.append(lem)
                arr.append(i)
            chunks.append(arr)
            meta.append((
                subj or "",
                int(grade) if pd.notna(grade) else None,
                int(gmax) if pd.notna(gmax) else None,
                doc,
            ))
        return {"vocab": vocab, "chunks": chunks, "meta": meta}

    def form_count(self, form: str) -> int:
        """Сколько раз словоформа встречается в учебниках всех предметов (для названий лемм)."""
        if not hasattr(self, "_forms"):
            cache = CACHE_DIR / "corpus_forms.pkl"
            if cache.exists() and cache.stat().st_mtime >= config.CORPUS_CHUNKS.stat().st_mtime:
                with open(cache, "rb") as f:
                    self._forms = pickle.load(f)
            else:
                self._forms = Counter()
                for t in self.texts():
                    self._forms.update(WORD_RE.findall(t.lower().replace("ё", "е")))
                with open(cache, "wb") as f:
                    pickle.dump(self._forms, f, protocol=pickle.HIGHEST_PROTOCOL)
        return self._forms.get(form, 0) or self._forms.get(form.rpartition("-")[2], 0) // 10

    def key_ids(self, key: str) -> tuple[int, ...] | None:
        out = []
        for w in key.split():
            i = self.ids.get(self.lemma(w))
            if i is None:
                return None
            out.append(i)
        return tuple(out)

    def counts(self, keys) -> dict[str, dict]:
        """
        {ключ: {'subj': Counter предмет -> вхождений, 'grade': Counter класс -> вхождений}};
        ключ — слова в начальной форме через пробел (термин из словаря или лемма).
        """
        targets = {k: self.key_ids(k) for k in keys}
        single = defaultdict(list)
        multi = defaultdict(list)
        for k, t in targets.items():
            if t is None:
                continue
            (single if len(t) == 1 else multi)[t[0]].append((k, t))
        res = {k: {"subj": Counter(), "grade": Counter()} for k in keys}
        for arr, (subj, grade, gmax, _) in zip(self.chunks, self.meta):
            c = Counter(arr)
            g = grade if grade is not None and gmax == grade else None
            for i in set(c) & single.keys():
                for k, _ in single[i]:
                    res[k]["subj"][subj] += c[i]
                    if g:
                        res[k]["grade"][g] += c[i]
            if multi:
                starts = set(c) & multi.keys()
                if not starts:
                    continue
                for pos, i in enumerate(arr):
                    if i in starts:
                        for k, t in multi[i]:
                            if tuple(arr[pos:pos + len(t)]) == t:
                                res[k]["subj"][subj] += 1
                                if g:
                                    res[k]["grade"][g] += 1
        return res

    def subject_profile(self, subj_counts: Counter) -> tuple[str, str]:
        """
        («да» / «нет», где встречается чаще всего). Частота — на миллион
        слов учебников предмета; маленькие предметы не учитываются.
        """
        ipm = {
            s: subj_counts.get(s, 0) / n * 1e6
            for s, n in self.subject_words.items()
            if n >= MIN_SUBJECT_WORDS or s in BIO_SUBJECTS
        }
        bio = sum(subj_counts.get(s, 0) for s in BIO_SUBJECTS) / sum(
            self.subject_words[s] for s in BIO_SUBJECTS) * 1e6
        others = {s: v for s, v in ipm.items() if s not in BIO_SUBJECTS and v > 0}
        top = sorted(others.items(), key=lambda x: -x[1])[:3]
        if not bio and not top:
            return "да", "в школьных учебниках не встречается"
        share = bio / (bio + sum(others.values()))
        where = (f"доля биологии {share:.0%}: " + "; ".join(
            [f"биология {bio:.0f}"] + [f"{s.lower()} {v:.0f}" for s, v in top]))
        if share >= BIO_SHARE:
            return "да", where
        spread = sum(1 for v in others.values() if v >= bio * 0.3)
        if spread >= 6:
            return "нет (общеупотребительное)", where
        return f"нет (чаще: {top[0][0].lower()})", where

    def texts(self) -> list[str]:
        if self._texts is None:
            self._texts = pd.read_parquet(config.CORPUS_CHUNKS, columns=["text"]).text.fillna("").tolist()
        return self._texts

    def find_definition(self, key: str, grade: int) -> str:
        """
        Определение слова в других учебниках корпуса: сначала биология,
        потом другие предметы; среди них — младшие классы раньше.
        """
        if len(key.replace(" ", "")) < 4:
            return ""
        target = self.key_ids(key)
        if target is None:
            return ""
        texts = self.texts()
        n = len(target)
        cands = []
        for ci, (arr, (subj, g, gmax, _)) in enumerate(zip(self.chunks, self.meta)):
            if target[0] not in arr:
                continue
            if n > 1 and not any(tuple(arr[p:p + n]) == target for p in range(len(arr) - n + 1)):
                continue
            cands.append((subj not in BIO_SUBJECTS, g or 99, ci))
        first_word = key.split()[0]
        for _, g, ci in sorted(cands)[:80]:
            text = texts[ci]
            for m in WORD_RE.finditer(text.lower()):
                if self.lemma(m.group()) != self.lemma(first_word):
                    continue
                end = m.end()
                if n > 1:
                    tail = WORD_RE.finditer(text.lower(), end)
                    ws = [next(tail, None) for _ in range(n - 1)]
                    if any(w is None for w in ws):
                        continue
                    if tuple(self.ids.get(self.lemma(w.group())) for w in ws) != target[1:]:
                        continue
                    end = ws[-1].end()
                ss, se = sentence_bounds(text, m.start(), end, para_breaks=True)
                if classify_definition(text[ss:se], m.start() - ss, end - ss) == "def":
                    subj = self.meta[ci][0]
                    grade_s = f"{g} кл." if g != 99 else ""
                    return f"{subj}, {grade_s}: {clean_sentence(with_context(text, ss, se, True), 350)}"
        return ""


# =============================================================================
# Упоминания в прошлых классах
# =============================================================================

class _Everything:
    """Множество «все слова знакомы»: при подсчёте терминов новые слова не ищутся."""
    def __contains__(self, item):
        return True


def textbook_term_counts(grade: int, pre, patterns, shadowed, use_cache: bool = True) -> Counter:
    """Сколько раз каждый термин словаря встретился в учебнике биологии grade класса."""
    import hashlib
    path = config.GRADE_DOCX[grade]
    digest = hashlib.md5("\n".join(p for v in patterns.values() for p in v).encode("utf-8")).hexdigest()[:10]
    cache = CACHE_DIR / f"textbook_{grade}_terms_{digest}.json"
    if use_cache and cache.exists() and cache.stat().st_mtime >= path.stat().st_mtime:
        return Counter(json.loads(cache.read_text(encoding="utf-8")))
    a = h.analyze_document(Document(str(path)), pre, _Everything(), patterns, shadowed, set())
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps(a.stats_terms, ensure_ascii=False), encoding="utf-8")
    return a.stats_terms


class PriorMentions:
    """
    Упоминания слова или термина в прошлых классах: в учебниках биологии
    5…N-1 классов и во всех учебниках 2…N-1 классов. Для класса берётся
    большее из чисел «учебник биологии» и «все учебники», как в базе
    знакомых слов.
    """

    def __init__(self, grade: int, pre, patterns, shadowed, use_cache: bool = True):
        self.grade = grade
        bio_grades = [g for g in sorted(config.GRADE_DOCX) if g < grade]
        self.bio_grades = bio_grades
        self.bio_terms = {g: textbook_term_counts(g, pre, patterns, shadowed, use_cache) for g in bio_grades}
        self.bio_lemmas = h.textbook_counts(bio_grades, pre, use_cache=use_cache)
        self.all_lemmas = h.rulex_counts([g for g in range(2, grade)], pre)

    def of(self, key: str, is_term: bool, members: list[str] | None = None,
           corpus_grades: Counter | None = None) -> dict:
        """corpus_grades — вхождения составного термина по классам во всех учебниках."""
        members = members or [key]
        multi = len(key.split()) > 1
        bio, allg = Counter(), Counter()
        for g in self.bio_grades:
            src = self.bio_terms[g] if is_term else self.bio_lemmas.get(g, Counter())
            bio[g] = sum(src.get(m, 0) for m in members)
        if multi:
            by_grade = corpus_grades or Counter()
            other = {g: by_grade.get(g, 0) for g in range(2, self.grade)}
        else:
            other = {g: sum(self.all_lemmas.get(g, Counter()).get(m, 0) for m in members) for g in range(2, self.grade)}
        for g in range(2, self.grade):
            allg[g] = max(other.get(g, 0), bio.get(g, 0))
        bio, allg = +bio, +allg
        last = max(allg) if allg else None
        return {"bio": bio, "all": allg, "last": last}

    @staticmethod
    def by_grade(c: Counter) -> str:
        return " · ".join(f"{g}: {int(n)}" for g, n in sorted(c.items())) or "—"

    def verdict(self, m: dict) -> str:
        total = int(sum(m["all"].values()))
        if not total:
            return "не встречалось"
        if total < config.MIN_KNOWN_FREQ:
            return f"редко: {total} {raz(total)}"
        if self.grade - m["last"] >= LONG_GAP_YEARS:
            return f"давно: последний раз в {m['last']} кл."
        if not m["bio"]:
            return "в биологии не было, встречалось в других предметах"
        last_bio = max(m["bio"])
        if self.grade - last_bio >= LONG_GAP_YEARS:
            return f"в биологии давно: последний раз в {last_bio} кл."
        return "встречалось недавно"


# =============================================================================
# Отчёт
# =============================================================================

def short_count(n) -> str:
    """Компактное число для пометки в тексте: 7, 693, 1,3т, 13т."""
    n = int(n)
    if n < 1000:
        return str(n)
    if n < 10000:
        return f"{n / 1000:.1f}".replace(".", ",").replace(",0", "") + "т"
    return f"{round(n / 1000)}т"


def grades_note(c: Counter) -> str:
    """«5:12 7:3 8:40» — упоминания по классам, классы без упоминаний опущены; «0» — ни разу."""
    parts = [f"{g}:{short_count(n)}" for g, n in sorted(c.items()) if n]
    return " ".join(parts) if parts else "0"


def term_annotations(a, use_cache: bool = True) -> dict:
    """
    Где в тексте поставить у термина числа упоминаний в прошлых классах
    по классам: сверху — учебники биологии («5:12 7:3»), снизу — все учебники.

    Чтобы не загромождать текст, числа ставятся при первом упоминании
    термина в учебнике и там, где термин снова появляется после перерыва
    в FAR_PAGES страниц и больше (такие места ученик мог забыть). Слово
    внутри более длинного термина («мембрана» в «клеточная мембрана»)
    отдельной пометки не получает.

    Returns:
        {"by_para": {абзац: [(позиция конца термина, верхнее, нижнее)]},
         "by_start": {начало вхождения в потоке: (ключ, верхнее, нижнее)},
         "bio_span", "all_span", "corpus", "prior", "pm"}
    """
    pdf = load_pdf(config.SOURCE_PDF[a.grade], use_cache)
    pm = PageMap(a, pdf)
    corpus = Corpus(a.preprocessor.morph, use_cache)
    sections, _ = parse_toc(pdf)
    last_section_page = max((sec["end"] for sec in sections), default=len(pdf["pages"]))

    def in_box(pi: int) -> bool:
        """Абзац в ячейке таблицы или в надписи: там поле с числами не помещается."""
        tags = {el.tag.rsplit("}", 1)[-1] for el in a.paragraphs[pi]._p.iterancestors()}
        return bool(tags & {"tc", "txbxContent"})

    occ = defaultdict(list)
    spans = []
    for kind, key, s, e in a.occurrences:
        if kind == "term":
            occ[key].append((s, e))
            spans.append((s, e))
    spans.sort()

    def inside_longer(s: int, e: int) -> bool:
        i = bisect_left(spans, (s, e))
        for s2, e2 in spans[max(0, i - 8):i + 8]:
            if s2 <= s and e <= e2 and (s2, e2) != (s, e):
                return True
        return False

    multi = [k for k in occ if " " in k]
    grade_counts = corpus.counts(multi) if multi else {}
    prior = PriorMentions(a.grade, a.preprocessor, a.patterns_by_len, a.verb_shadowed, use_cache)
    by_para, by_start = defaultdict(list), {}
    for key, lst in occ.items():
        m = prior.of(key, True, corpus_grades=grade_counts.get(key, {}).get("grade"))
        sup, sub = grades_note(m["bio"]), grades_note(m["all"])
        last_page = None
        for s, e in sorted(lst):
            page = pm.page_of(s)
            # Приложения и указатель терминов после последнего параграфа не размечаются
            if inside_longer(s, e) or (page and page > last_section_page):
                continue
            if last_page is None or (page and page - last_page >= FAR_PAGES):
                rng = a.stream.to_paragraphs(max(s, e - 1), e)
                if rng and not in_box(rng[-1][0]):
                    pi, _, pe = rng[-1]
                    by_para[pi].append((pe, sup, sub))
                    by_start[s] = (key, sup, sub)
            if page:
                last_page = page
    bio_grades = prior.bio_grades
    return {
        "by_para": by_para, "by_start": by_start, "corpus": corpus, "prior": prior, "pm": pm,
        "bio_span": f"{bio_grades[0]}–{bio_grades[-1]}" if bio_grades else "",
        "all_span": f"2–{a.grade - 1}",
    }


def distances(idx: list[int], pages: list[int | None]) -> tuple[int | str, str, int]:
    """(среднее расстояние, наибольший разрыв «N (с. A → с. B)», разрыв в страницах)."""
    if len(idx) < 2:
        return "", "", 0
    gaps = [b - a for a, b in zip(idx, idx[1:])]
    k = max(range(len(gaps)), key=gaps.__getitem__)
    pa, pb = pages[k], pages[k + 1]
    far = (pb - pa) if pa and pb else 0
    return round(statistics.mean(gaps)), f"{gaps[k]} (с. {pa} → с. {pb})", far


def collect(a, pdf: dict, corpus: Corpus, use_cache: bool = True) -> dict:
    """Всё, что нужно для Excel и PDF-отчёта."""
    pre = a.preprocessor
    morph = pre.morph
    pm = PageMap(a, pdf)
    starts = word_starts(a)
    widx = lambda pos: bisect_left(starts, pos)
    sections = find_sections(a, pm, pdf)

    occ = defaultdict(list)          # (вид, ключ) -> [(start, end)]
    for kind, key, s, e in a.occurrences:
        occ[(kind, key)].append((s, e))
    surfaces = defaultdict(Counter)
    for (kind, key), lst in occ.items():
        if kind == "new":
            for s, e in lst:
                surfaces[key][a.stream.text[s:e].lower()] += 1
    canon = merge_variants(surfaces, a.stats_new, morph, corpus.form_count)
    groups = defaultdict(list)
    for l, c in canon.items():
        groups[c].append(l)
    new_occ = {c: sorted(x for m in ms for x in occ[("new", m)]) for c, ms in groups.items()}
    new_forms = {c: sum((surfaces[m] for m in ms), Counter()) for c, ms in groups.items()}
    term_occ = {key: sorted(lst) for (kind, key), lst in occ.items() if kind == "term"}
    other_occ = {(kind, key): lst for (kind, key), lst in occ.items() if kind in ("skipped", "derived")}

    def info(lst):
        idx = [widx(s) for s, _ in lst]
        pages = [pm.page_of(s) for s, _ in lst]
        avg, gap, far = distances(idx, pages)
        return {"pages": pages, "idx": idx, "avg": avg, "gap": gap, "far": far}

    new_info = {k: info(v) for k, v in new_occ.items()}
    term_info = {k: info(v) for k, v in term_occ.items()}
    other_info = {k: info(v) for k, v in other_occ.items()}

    print("[Отчёт] определения в учебнике…")
    # Определение ищется у всех вхождений слова, а не только выделенных:
    # первое упоминание могло попасть в термин или в рубрику без подсветки
    lemma_pos = defaultdict(list)
    for b_start, _, tokens in a.blocks:
        for w in build_indexed_words(tokens):
            lemma_pos[w["lemma"]].append((b_start + w["start"], b_start + w["end"]))
    new_all = {c: sorted({x for m in ms for x in lemma_pos.get(m, [])} | set(new_occ[c])) for c, ms in groups.items()}
    new_defs = find_definitions(a, new_all, pm)
    term_defs = find_definitions(a, term_occ, pm)

    print("[Отчёт] предметы и упоминания в прошлых классах…")
    subj = corpus.counts(list(new_occ) + list(term_occ))
    prior = PriorMentions(a.grade, pre, a.patterns_by_len, a.verb_shadowed, use_cache)
    new_prior = {k: prior.of(k, False, groups[k]) for k in new_occ}
    term_prior = {k: prior.of(k, True, corpus_grades=subj[k]["grade"]) for k in term_occ}

    print("[Отчёт] определения в других учебниках…")
    other_defs = {}
    for k, d in list(new_defs.items()) + [(t, d) for t, d in term_defs.items() if t in a.new_terms]:
        if d["status"] != "есть":
            other_defs[k] = corpus.find_definition(k, a.grade)

    # Слова по параграфам и страницам
    word_sec = []
    sec_words, page_words = Counter(), Counter()
    for pos in starts:
        pi = pm.para_of(pos)
        num, main = section_of_para(sections, pm, pi)
        word_sec.append(num if main else None)
        if main:
            sec_words[num] += 1
            page_words[pm.para_page[pi]] += 1

    return {
        "pm": pm, "sections": sections, "word_sec": word_sec, "sec_words": sec_words, "page_words": page_words,
        "canon": canon, "groups": groups, "new_occ": new_occ, "new_forms": new_forms, "term_occ": term_occ,
        "new_info": new_info, "term_info": term_info, "other_info": other_info,
        "new_defs": new_defs, "term_defs": term_defs, "other_defs": other_defs,
        "subj": subj, "prior": prior, "new_prior": new_prior, "term_prior": term_prior,
        "corpus": corpus,
    }


def section_rows(a, R: dict) -> list[list]:
    """Строки вкладки «Параграфы»."""
    word_sec = R["word_sec"]
    terms_by_sec, new_by_sec = defaultdict(Counter), defaultdict(Counter)
    new_idx_by_sec = defaultdict(list)
    for key, inf in R["term_info"].items():
        for i in inf["idx"]:
            if i < len(word_sec) and word_sec[i]:
                terms_by_sec[word_sec[i]][key] += 1
    for key, inf in R["new_info"].items():
        for i in inf["idx"]:
            if i < len(word_sec) and word_sec[i]:
                new_by_sec[word_sec[i]][key] += 1
                new_idx_by_sec[word_sec[i]].append(i)
    dens = {}
    for s in R["sections"]:
        n = R["sec_words"][s["num"]]
        dens[s["num"]] = sum(terms_by_sec[s["num"]].values()) / n * 100 if n else 0
    vals = sorted(v for v in dens.values() if v)
    q1 = vals[len(vals) // 4] if vals else 0
    q3 = vals[(3 * len(vals)) // 4] if vals else 0
    rows = []
    for s in R["sections"]:
        num = s["num"]
        n = R["sec_words"][num]
        t, nw = terms_by_sec[num], new_by_sec[num]
        d = dens[num]
        level = "высокая" if d >= q3 else ("низкая" if d <= q1 else "средняя")
        idx = sorted(new_idx_by_sec[num])
        gap = round(statistics.mean(b - a_ for a_, b in zip(idx, idx[1:]))) if len(idx) > 1 else ""
        main_end = (s["trainer_page"] or s["end"])
        rows.append([
            f"§ {num}", s["title"], f"{s['start']}–{s['end']}",
            f"{s['start']}–{main_end}" + ("" if s["trainer_page"] else " (рубрики «Тренируемся» нет)"),
            n, sum(t.values()), len(t), round(d, 1), level,
            sum(nw.values()), len(nw), round(sum(nw.values()) / n * 1000, 1) if n else 0, gap,
            ", ".join(f"{k} ({c})" for k, c in t.most_common(5)),
            ", ".join(k for k, _ in nw.most_common(8)),
        ])
    R["section_density"] = dens
    R["density_q3"] = q3
    return rows


SECTION_COLUMNS = [
    "Параграф", "Название", "Страницы параграфа", "Страницы основного текста (до «Тренируемся»)",
    "Слов в основном тексте", "Вхождений терминов", "Разных терминов", "Терминов на 100 слов",
    "Плотность терминов", "Вхождений новых слов", "Разных новых слов", "Новых слов на 1000 слов",
    "Среднее расстояние между новыми словами, слов", "Самые частые термины", "Новые слова параграфа",
]


def write_excel_report(a, out_path: Path, pdf_path: Path | None = None, corpus: Corpus | None = None,
                       use_cache: bool = True) -> dict:
    """
    Пишет xlsx-отчёт. Возвращает собранные данные (для PDF-отчёта).
    """
    pdf = load_pdf(pdf_path or config.SOURCE_PDF[a.grade], use_cache)
    corpus = corpus or Corpus(a.preprocessor.morph, use_cache)
    R = collect(a, pdf, corpus, use_cache)
    prior: PriorMentions = R["prior"]
    morph = a.preprocessor.morph
    gr = a.grade
    bio_span = f"{prior.bio_grades[0]}–{prior.bio_grades[-1]}" if prior.bio_grades else "—"
    all_span = f"2–{gr - 1}"
    BIO_COL = f"Упоминаний раньше: учебники биологии {bio_span} кл."
    ALL_COL = f"Упоминаний раньше: все учебники {all_span} кл."
    BY_COL = "По классам (класс: упоминаний; сверху — учебники биологии, снизу — все учебники)"

    def prior_cells(m):
        return [int(sum(m["bio"].values())), int(sum(m["all"].values())),
                f"{prior.by_grade(m['bio'])}\n{prior.by_grade(m['all'])}"]

    # --- Новые слова -----------------------------------------------------------
    new_rows = []
    for k in sorted(R["new_occ"], key=lambda k: (-len(R["new_occ"][k]), k)):
        inf, d, m = R["new_info"][k], R["new_defs"][k], R["new_prior"][k]
        bio, where = corpus.subject_profile(R["subj"][k]["subj"])
        forms = R["new_forms"][k]
        new_rows.append([
            k, ", ".join(f"{f} ({n})" for f, n in forms.most_common()), len(R["new_occ"][k]),
            pages_str(inf["pages"]), inf["avg"], inf["gap"],
            d["status"], (f"с. {d['page']}: {d['text']}" if d["text"] else ""),
            R["other_defs"].get(k, ""), "; ".join(d["figures"]), bio, where, *prior_cells(m),
            "да" if any(morph.word_is_known(x) for x in R["groups"][k]) else "",
        ])
    new_df = pd.DataFrame(new_rows, columns=[
        "Новое слово (лемма)", "Формы в тексте (вхождений)", "Вхождений в учебнике", "Страницы",
        "Среднее расстояние между вхождениями, слов", "Наибольший разрыв между вхождениями, слов (страницы)",
        "Определение в учебнике", "Текст определения (страница)", "Определение в других учебниках",
        "Рисунок (где слово показано или упомянуто рядом)", "Биологическое слово", "Частота по предметам (на млн слов)", BIO_COL, ALL_COL, BY_COL, h.DICT_COLUMN,
    ])

    # --- Термины ---------------------------------------------------------------
    term_rows = []
    for k in sorted(R["term_occ"], key=lambda k: (-len(R["term_occ"][k]), k)):
        inf, d, m = R["term_info"][k], R["term_defs"][k], R["term_prior"][k]
        bio, where = corpus.subject_profile(R["subj"][k]["subj"])
        term_rows.append([
            k, len(R["term_occ"][k]), pages_str(inf["pages"]), inf["avg"], inf["gap"],
            "да" if k in a.new_terms else "", dict_names(a, k), bio, where,
            d["status"], (f"с. {d['page']}: {d['text']}" if d["text"] else ""),
            R["other_defs"].get(k, "") if k in a.new_terms else "", "; ".join(d["figures"]),
            *prior_cells(m), m["last"] or "", prior.verdict(m),
        ])
    terms_df = pd.DataFrame(term_rows, columns=[
        "Термин (слова в начальной форме)", "Вхождений в учебнике", "Страницы",
        "Среднее расстояние между вхождениями, слов", "Наибольший разрыв между вхождениями, слов (страницы)",
        "Есть незнакомое ученику слово", "Словари", "Биологический термин", "Частота по предметам (на млн слов)",
        "Определение в учебнике", "Текст определения (страница)", "Определение в других учебниках",
        "Рисунок (где слово показано или упомянуто рядом)",
        BIO_COL, ALL_COL, BY_COL, "Последний раз встречался, класс", "Помнит ли ученик",
    ])

    # --- Параграфы -------------------------------------------------------------
    sec_df = pd.DataFrame(section_rows(a, R), columns=SECTION_COLUMNS)

    # --- Не выделено ---------------------------------------------------------------
    skipped_rows = []
    for word, count in sorted(a.stats_skipped.items(), key=lambda x: (x[0] not in a.junk, -x[1], x[0])):
        inf = R["other_info"].get(("skipped", word), {"pages": []})
        skipped_rows.append([word, count, pages_str(inf["pages"]),
                             "мусор распознавания" if word in a.junk else "редкое, но знакомое ученику"])
    for word, (count, base) in sorted(a.stats_derived.items()):
        inf = R["other_info"].get(("derived", word), {"pages": []})
        skipped_rows.append([word, count, pages_str(inf["pages"]), f"образовано от знакомого: {base}"])
    skipped_df = pd.DataFrame(skipped_rows, columns=["Лемма", "Вхождений в учебнике", "Страницы", "Причина"])

    # --- Сводка ---------------------------------------------------------------------
    all_new_idx = sorted(i for inf in R["new_info"].values() for i in inf["idx"])
    avg_new_gap = round(statistics.mean(b - x for x, b in zip(all_new_idx, all_new_idx[1:]))) if len(all_new_idx) > 1 else ""
    n_words = len(word_starts(a))
    defs = Counter(R["new_defs"][k]["status"] for k in R["new_occ"])
    high = [f"§ {s['num']}" for s in R["sections"] if R["section_density"][s["num"]] >= R["density_q3"]]
    min_freq = a.min_freq
    summary = [
        ["Слово считается новым, если",
         f"в предыдущих классах оно встретилось меньше {min_freq} раз" if min_freq > 1 else "в предыдущих классах его нет"],
        ["Где считались упоминания", a.sources],
        [f"Знакомых слов (упоминаний: {min_freq} и больше)", len(a.known)],
        ["Слов в учебнике", n_words],
        ["Новых слов (разных)", len(R["new_occ"])],
        ["Новых слов (всего вхождений в учебнике)", sum(len(v) for v in R["new_occ"].values())],
        ["  из них с определением в учебнике", defs["есть"]],
        ["  только с пояснением", defs["только пояснение"]],
        ["  без определения, но показаны на рисунке", defs[FIG_STATUS]],
        ["  выделены жирным, но без определения", defs["выделено жирным, определения нет"]],
        ["  без определения", defs["нет"]],
        ["Новых слов, которые показаны на рисунке или упомянуты рядом с ним",
         sum(1 for k in R["new_occ"] if R["new_defs"][k]["figures"])],
        ["Среднее расстояние между соседними новыми словами в тексте, слов", avg_new_gap],
        ["Терминов найдено (разных)", len(a.stats_terms)],
        ["  из них с незнакомым ученику словом", len(a.new_terms)],
        ["  из них биологических", int((terms_df["Биологический термин"] == "да").sum())],
        ["Терминов найдено (всего вхождений в учебнике)", sum(a.stats_terms.values())],
        ["Терминов в словарях (после приведения к начальной форме)", a.patterns_total],
        *dictionary_summary(a),
        ["Параграфов", len(R["sections"])],
        ["Параграфы с высокой плотностью терминов", ", ".join(high)],
        ["Не выделено: мусор распознавания (лемм)", sum(1 for w in a.stats_skipped if w in a.junk)],
        ["Не выделено: редкие, но знакомые слова (лемм)", sum(1 for w in a.stats_skipped if w not in a.junk)],
        ["Не выделено: образованы от знакомых слов (лемм)", len(a.stats_derived)],
    ]
    summary_df = pd.DataFrame(summary, columns=["Показатель", "Значение"])

    # --- Проверка определений: все предложения со словами без определения ---------
    check_rows = []
    items = [("новое слово", k, R["new_defs"][k]) for k in new_df["Новое слово (лемма)"]]
    items += [("термин с незнакомым словом", k, R["term_defs"][k])
              for k in terms_df["Термин (слова в начальной форме)"] if k in a.new_terms]
    for kind, k, d in items:
        if d["status"] == "есть":
            continue
        figs = "; ".join(d["figures"])
        for page, sent, ws, we in d["contexts"]:
            check_rows.append([k, kind, d["status"], figs, page or "", sent, (ws, we)])
    R["check_rows"] = check_rows
    guide_df = pd.DataFrame(guide_rows(a, bio_span, all_span, morph), columns=["Вкладка / столбец", "Что это"])

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(out_path), engine="openpyxl") as w:
        summary_df.to_excel(w, sheet_name="Сводка", index=False)
        new_df.to_excel(w, sheet_name="Новые слова", index=False)
        terms_df.to_excel(w, sheet_name="Термины", index=False)
        sec_df.to_excel(w, sheet_name="Параграфы", index=False)
        dictionaries_df(a, R).to_excel(w, sheet_name="Словари", index=False)
        skipped_df.to_excel(w, sheet_name="Не выделено", index=False)
        pd.DataFrame([r[:6] for r in check_rows], columns=CHECK_COLUMNS).to_excel(
            w, sheet_name="Проверка определений", index=False)
        guide_df.to_excel(w, sheet_name="Пояснения", index=False)
    format_workbook(out_path, check_rows)
    print(f"[Отчёт] Сохранён: {out_path}")
    R.update(new_df=new_df, terms_df=terms_df, sec_df=sec_df)
    return R


def dict_names(a, key: str) -> str:
    """«ФРП, указатель учебника, словарь терминов» — в каких словарях термин."""
    import term_dictionaries as td
    srcs = (a.get("term_sources") or {}).get(key, frozenset({"dict"}))
    return ", ".join(td.SOURCE_NAMES[s] for s in td.SOURCES if s in srcs)


def dictionary_summary(a) -> list[list]:
    import term_dictionaries as td
    sources = a.get("term_sources") or {}
    found = a.stats_terms
    rows = []
    for src in td.SOURCES:
        pats = [p for p, ss in sources.items() if src in ss]
        if not pats:
            continue
        hit = [p for p in pats if found.get(p)]
        rows.append([f"Словарь «{td.SOURCE_NAMES[src]}»: терминов / найдено в учебнике", f"{len(pats)} / {len(hit)}"])
    return rows


def dictionaries_df(a, R) -> pd.DataFrame:
    """
    Термины указателя и ФРП: встречается ли термин в учебнике, сколько раз,
    на каких страницах и в каких ещё словарях он есть.
    """
    import term_dictionaries as td
    sources = a.get("term_sources") or {}
    term_of = a.get("term_of") or {}
    rows = []
    for src in ("frp", "index"):
        for (s, pat), term in term_of.items():
            if s != src or term in td.SEARCH_ONLY:
                continue
            occ = R["term_info"].get(pat)
            n = len(R["term_occ"].get(pat, []))
            others = ", ".join(td.SOURCE_NAMES[o] for o in td.SOURCES if o in sources.get(pat, ()) and o != src)
            rows.append([td.SOURCE_NAMES[src], term, n, pages_str(occ["pages"]) if occ else "",
                         "" if n else "в учебнике не встречается", others])
    df = pd.DataFrame(rows, columns=["Словарь", "Термин", "Вхождений в учебнике", "Страницы", "Замечание",
                                     "Есть также в словарях"])
    return df.sort_values(["Словарь", "Вхождений в учебнике", "Термин"], ascending=[True, True, True])


def guide_rows(a, bio_span: str, all_span: str, morph) -> list[list[str]]:
    return [
        ["Сводка", "Итоговые числа по учебнику."],
        ["Новые слова", "Слова, выделенные бирюзовым: в прошлых классах встречались меньше "
                        f"{a.min_freq} раз. Разные падежные формы одного слова собраны в одну строку."],
        ["Термины", "Термины из трёх словарей: общего словаря терминов (жёлтый), указателя терминов в конце "
                    "учебника (зелёный) и ФРП — федеральной рабочей программы, то есть того, что требуется от "
                    "ученика (розовый). Если термин в нескольких словарях, фон — по старшему словарю (ФРП, затем "
                    "указатель, затем словарь), подчёркивание — цветом второго словаря, двойное — во всех трёх. "
                    "Столбец «Словари» перечисляет, где термин есть."],
        ["Словари", "Все термины указателя и ФРП: сколько раз и на каких страницах термин встречается в "
                    "учебнике. «в учебнике не встречается» — термин из словаря в тексте не найден (для ФРП это "
                    "значит, что требуемое программой понятие в учебнике не употребляется — или употребляется в "
                    "другой форме, которую стоит проверить)."],
        ["Параграфы", "Плотность терминов и новых слов по параграфам (см. ниже, как определены границы)."],
        ["Не выделено", "Слова, которые формально редкие, но не выделены: мусор распознавания, "
                        "знакомые ученику слова, слова, образованные от знакомых."],
        ["Проверка определений", "Для каждого нового слова и термина с незнакомым словом, у которого "
                                 "в учебнике не найдено полного определения, — все предложения, где оно "
                                 "встречается, со страницами. Само слово выделено жирным. По ним можно "
                                 "убедиться, что определения действительно нет. Одинаковые предложения "
                                 "(подписи к рисункам дублируются при конвертации) показаны один раз."],
        ["Страницы", "Страницы печатного учебника (по PDF издательства), на которых встречается слово."],
        ["Среднее расстояние между вхождениями", "Сколько слов текста в среднем проходит между соседними "
                                                 "вхождениями этого слова. Пусто — слово встречается один раз."],
        ["Наибольший разрыв", "Самый длинный промежуток между соседними вхождениями (в словах) и страницы, "
                              "между которыми он приходится. Большой разрыв — ученик мог забыть слово к "
                              "следующей встрече."],
        ["Определение в учебнике", "«есть» — найдено предложение-определение: «X — это…», «X — …» "
                                   "в начале предложения, «…называют X», «X называется / представляет "
                                   "собой…», «под X понимают…»; «только пояснение» — «X, т. е. …», "
                                   "«X, или Y», «X (от греч. …)»; «нет определения, есть рисунок» — отдельный "
                                   "случай: фразы-определения нет, но слово стоит в подписи к рисунку "
                                   "или предложение ссылается на рисунок; «выделено жирным, определения нет» — "
                                   "слово набрано жирным, но определяющей фразы рядом нет; «нет» — "
                                   "ничего из этого. Определение ищется автоматически по шаблонам, "
                                   "поэтому его стоит проверять глазами."],
        ["Рисунок", "Номер рисунка и страница, если слово стоит в подписи к рисунку или в предложении "
                    "со ссылкой на рисунок («(рис. 34)»). «рисунок» без номера — подпись или схема, номер "
                    "которой при конвертации потерялся."],
        ["Определение в других учебниках", "Для слов без определения: предложение-определение из "
                                           "школьного учебника другого класса или предмета (сначала "
                                           "биология). Можно взять за образец."],
        ["Биологическое слово / термин", "Сколько раз слово встречается на миллион слов в школьных "
                                         "учебниках каждого предмета (2–11 классы). «да» — на биологию "
                                         f"приходится не меньше {BIO_SHARE:.0%} суммы этих частот или слово "
                                         "есть только в биологии (в учебниках других предметов его нет); "
                                         "«нет (чаще: химия)» — слово в основном из другого предмета; "
                                         "«нет (общеупотребительное)» — встречается во многих предметах. "
                                         "Оценка статистическая: у слов с несколькими значениями («ядро», "
                                         "«ткань») частоты складываются по всем значениям."],
        ["Упоминаний раньше", f"Сколько раз слово встретилось в прошлых классах: в учебниках биологии "
                              f"{bio_span} классов и во всех школьных учебниках {all_span} классов. В "
                              f"столбце «По классам» верхняя строка — учебники биологии, нижняя — все "
                              f"учебники; «7: 12» значит «в 7 классе 12 раз»."],
        ["Помнит ли ученик", f"«давно» — слово не встречалось {LONG_GAP_YEARS} года и больше, его стоит "
                             f"напомнить; «редко» — встречалось меньше {a.min_freq} раз; «не встречалось» — "
                             f"в прошлых классах его не было."],
        ["Границы параграфа", "Параграф начинается на странице из оглавления и идёт до начала следующего "
                              "параграфа или главы. В плотность входит только основной текст: от начала "
                              "параграфа до рубрики «Тренируемся». Сама рубрика и всё, что после неё "
                              "(вопросы, «Совершенствуемся», «Обсуждаем», «Это интересно», «Шаги к "
                              "успеху»), не считаются. Страницы глав, введения и приложений не входят."],
        ["Плотность терминов", "Вхождений терминов на 100 слов основного текста. «высокая» — четверть "
                               "параграфов с самой большой плотностью, «низкая» — с самой маленькой."],
        [h.DICT_COLUMN, h.dictionary_note(morph)],
    ]


CHECK_COLUMNS = ["Слово / термин", "Что это", "Определение в учебнике", "Рисунок", "Страница",
                 "Предложение из учебника (слово выделено жирным)"]


def format_workbook(path: Path, check_rows: list | None = None) -> None:
    wb = load_workbook(path)
    if check_rows and "Проверка определений" in wb.sheetnames:
        # Слово в предложении — жирным, чтобы его было легко найти глазами
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        ws = wb["Проверка определений"]
        bold = InlineFont(b=True, color="C00000")
        for i, row in enumerate(check_rows, start=2):
            sent, (a_, b_) = row[5], row[6]
            if 0 <= a_ < b_ <= len(sent):
                parts = [p for p in (sent[:a_], TextBlock(bold, sent[a_:b_]), sent[b_:]) if p]
                ws.cell(row=i, column=6).value = CellRichText(*parts)
    widths = {
        "Сводка": [60, 70],
        "Новые слова": [24, 34, 11, 22, 14, 22, 20, 70, 70, 20, 14, 34, 13, 13, 30, 14],
        "Термины": [32, 11, 30, 14, 22, 13, 26, 12, 34, 20, 70, 60, 20, 13, 13, 30, 13, 22],
        "Параграфы": [9, 40, 12, 18, 11, 11, 10, 10, 11, 11, 10, 11, 14, 50, 50],
        "Не выделено": [26, 12, 22, 40],
        "Проверка определений": [26, 18, 24, 18, 9, 130],
        "Словари": [20, 42, 12, 30, 28, 34],
        "Пояснения": [36, 120],
    }
    wrap_cols = {
        "Новые слова": {2, 8, 9, 10, 12, 15}, "Термины": {3, 7, 9, 11, 12, 13, 16}, "Словари": {2, 4, 6}, "Параграфы": {2, 14, 15},
        "Пояснения": {2}, "Сводка": {2}, "Проверка определений": {4, 6},
    }
    for ws in wb.worksheets:
        ws.freeze_panes = "B2"
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
        for i, wdt in enumerate(widths.get(ws.title, []), 1):
            ws.column_dimensions[get_column_letter(i)].width = wdt
        cols = wrap_cols.get(ws.title, set())
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if c.column in cols:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
                else:
                    c.alignment = Alignment(vertical="top")
        ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# =============================================================================
# Запуск без пересборки docx
# =============================================================================

def analyze(grade: int, min_freq: int, use_cache: bool = True):
    from src.preprocessing import TextPreprocessorWithoutStopwords
    pre = TextPreprocessorWithoutStopwords(stopwords_path=config.STOPWORDS_PATH,
                                           replacements_path=config.REPLACEMENTS_PATH)
    known, sources = h.known_lemmas(grade, "both", pre, use_cache=use_cache, min_freq=min_freq)
    junk = h.load_not_new_words(config.NOT_NEW_WORDS_PATH)
    familiar = h.load_not_new_words(config.FAMILIAR_WORDS_PATH)
    import term_dictionaries
    patterns, shadowed, term_sources, term_lists, term_of = term_dictionaries.load_all_patterns(grade, pre)
    in_path = config.SOURCE_DOCX[grade]
    a = h.analyze_document(Document(str(in_path)), pre, known, patterns, shadowed, junk | familiar,
                           term_sources=term_sources)
    a.update(
        in_path=Path(in_path), grade=grade, min_freq=min_freq, known=known, sources=sources, junk=junk,
        patterns_by_len=patterns, patterns_total=sum(len(v) for v in patterns.values()), preprocessor=pre,
        verb_shadowed=shadowed, term_sources=term_sources, term_lists=term_lists, term_of=term_of,
        new_terms={t for t in a.stats_terms if any(not h.is_familiar(l, known, pre) for l in t.split())},
    )
    return a


def report_path(a, suffix: str) -> Path:
    return h.HIGHLIGHT_DIR / f"Порог {a.min_freq}" / f"{h._output_stem(a.in_path)}_{suffix}"


def main() -> None:
    ap = argparse.ArgumentParser(description="Подробный отчёт по учебнику (Excel и PDF-черновик)")
    ap.add_argument("--grade", type=int, required=True)
    ap.add_argument("--min-freq", type=int, default=config.MIN_KNOWN_FREQ)
    ap.add_argument("--pdf", action="store_true", help="сделать и PDF-отчёт для издателя")
    ap.add_argument("--no-cache", action="store_true")
    args = ap.parse_args()
    a = analyze(args.grade, args.min_freq, not args.no_cache)
    R = write_excel_report(a, report_path(a, f"отчёт_порог{a.min_freq}.xlsx"), use_cache=not args.no_cache)
    if args.pdf:
        import publisher_pdf
        publisher_pdf.write_pdf_report(a, R, report_path(a, f"рекомендации_издателю_порог{a.min_freq}.pdf"))


if __name__ == "__main__":
    main()
