"""
Сравнение двух готовых версий разметки: порог 1 и порог 3.

Обе версии строятся одинаково (чистка мусора распознавания и список знакомых
слов включены), различается только порог: сколько раз слово должно встретиться
в прошлых классах, чтобы считаться знакомым ученику.

Запуск: python compare_thresholds.py [выходной xlsx]
"""
from __future__ import annotations

import csv
import sys
from collections import Counter, defaultdict

import pandas as pd
from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import config
import highlight_new_and_terms as h
import textbook_report as tr
from src.preprocessing import TextPreprocessorWithoutStopwords

OUT = sys.argv[1] if len(sys.argv) > 1 else str(h.HIGHLIGHT_DIR / "Сравнение_порог1_и_порог3.xlsx")
THRESHOLDS = (1, 3)


def raz(n: int) -> str:
    """Склонение «раз»: 1 раз, 2 раза, 5 раз, 21 раз, 22 раза."""
    return "раза" if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14) else "раз"


def grouped_new(a, pm, morph):
    """Новые слова с объединёнными вариантами леммы: {лемма: (вхождений, страницы, пример, формы)}."""
    occ = defaultdict(list)
    for kind, key, s, e in a.occurrences:
        if kind == "new":
            occ[key].append((s, e))
    surfaces = {k: Counter(a.stream.text[s:e] for s, e in v) for k, v in occ.items()}
    canon = tr.merge_variants(surfaces, a.stats_new, morph, corpus.form_count)
    out = {}
    for lemma, c in canon.items():
        n, pages, ex, forms = out.get(c, (0, [], "", Counter()))
        lst = occ[lemma]
        pages = pages + [pm.page_of(s) for s, _ in lst]
        if not ex:
            s, e = lst[0]
            ex = a.stream.text[max(0, s - 50):e + 50].replace("\n", " / ")
        out[c] = (n + len(lst), pages, ex, forms + surfaces[lemma])
    return out, canon


pre = TextPreprocessorWithoutStopwords(stopwords_path=config.STOPWORDS_PATH, replacements_path=config.REPLACEMENTS_PATH)
morph = pre.morph
patterns, shadowed = h.load_patterns(config.DICT_PATH, pre)
corpus = tr.Corpus(morph)
excl = h.load_not_new_words(config.NOT_NEW_WORDS_PATH) | h.load_not_new_words(config.FAMILIAR_WORDS_PATH)

sharov = {}
with open(config.RULEX_DATA_DIR / "sharov.csv", encoding="utf-8") as f:
    rd = csv.reader(f, delimiter="\t")
    next(rd)
    for row in rd:
        if len(row) > 2:
            w = row[0].strip().lower().replace("ё", "е")
            try:
                sharov[w] = max(sharov.get(w, 0), float(row[2]))
            except ValueError:
                pass

summary, sheets_words, term_rows = [], {}, []

for grade, path in config.SOURCE_DOCX.items():
    prev_grades = [g for g in range(2, 12) if g < grade]
    book_grades = [g for g in sorted(config.GRADE_DOCX) if g < grade]
    rulex = h.rulex_counts(prev_grades, pre)
    books = h.textbook_counts(book_grades, pre)

    per_grade = defaultdict(list)
    for g, c in list(rulex.items()) + list(books.items()):
        per_grade[g].append(c)
    total, by_grade = Counter(), defaultdict(dict)
    for g, counters in per_grade.items():
        for w in set().union(*counters):
            n = max(c.get(w, 0) for c in counters)
            if n:
                total[w] += n
                by_grade[w][g] = n

    doc = Document(str(path))
    results, base = {}, None
    for th in THRESHOLDS:
        known, _ = h.known_lemmas(grade, "both", pre, min_freq=th)
        a = h.analyze_document(doc, pre, known, patterns, shadowed, excl, pretokenized=base)
        base = base or a
        new_terms = {t for t in a.stats_terms if any(not h.is_familiar(w, known, pre) for w in t.split())}
        results[th] = (known, a, new_terms)
    pm = tr.PageMap(base, tr.load_pdf(config.SOURCE_PDF[grade]))
    (known1, a1, nt1), (known3, a3, nt3) = results[1], results[3]
    new1, canon1 = grouped_new(a1, pm, morph)
    new3, canon3 = grouped_new(a3, pm, morph)

    stopped = sorted(set(new1) - set(new3))
    assert not stopped, f"при пороге 3 перестали выделяться слова: {stopped}"
    added = sorted(set(new3) - set(new1))

    def mentions(lemma, surf):
        """Сколько раз слово встретилось в прошлых классах (по лемме или по форме pymorphy3)."""
        keys = [lemma] + [m for m, c in canon3.items() if c == lemma] + [
            p.normal_form.replace("ё", "е") for s in surf for p in morph.parse(s.lower())]
        for k in dict.fromkeys(keys):
            if total.get(k):
                grades = ", ".join(f"{g} кл. — {n:g}" for g, n in sorted(by_grade[k].items()))
                return int(total[k]), (grades if k == lemma else f"как «{k}»: {grades}")
        return 0, ""

    rows = []
    for lemma in added:
        n_occ, pages, ex, forms = new3[lemma]
        n, grades = mentions(lemma, list(forms))
        if n:
            why = f"мало упоминаний: встретилось {n} {raz(n)}, а для «знакомого» нужно 3"
        else:
            why = "мало упоминаний: само слово и его части встречались реже 3 раз"
        rows.append([lemma, n_occ, tr.pages_str(pages), n, grades, why, sharov.get(lemma, 0) or "",
                     "да" if morph.word_is_known(lemma) else "", ex])
    df = pd.DataFrame(rows, columns=[
        "Слово (лемма)", "Вхождений в учебнике", "Страницы", "Упоминаний в прошлых классах",
        "По классам (упоминаний)", "Почему стало выделяться", "Частота в языке (словарь Шарова, на млн слов)",
        h.DICT_COLUMN, "Пример из учебника",
    ]).sort_values(["Упоминаний в прошлых классах", "Вхождений в учебнике"], ascending=[False, False])
    sheets_words[grade] = df

    # Термины, ставшие «с незнакомым словом» из-за порога 3
    term_pages = defaultdict(list)
    for kind, key, s, e in a3.occurrences:
        if kind == "term":
            term_pages[key].append(pm.page_of(s))
    for term in sorted(nt3 - nt1):
        unfamiliar = [w for w in term.split() if not h.is_familiar(w, known3, pre)]
        n = sum(total.get(w, 0) for w in unfamiliar)
        term_rows.append([grade, term, ", ".join(unfamiliar), n, a3.stats_terms[term], tr.pages_str(term_pages[term])])

    occ1, occ3 = sum(v[0] for v in new1.values()), sum(v[0] for v in new3.values())
    summary += [
        [grade, "Новых слов (разных)", len(new1), len(new3), len(new3) - len(new1)],
        [grade, "Новых слов (всего вхождений в учебнике)", occ1, occ3, occ3 - occ1],
        [grade, "Терминов (разных)", len(a1.stats_terms), len(a3.stats_terms), len(a3.stats_terms) - len(a1.stats_terms)],
        [grade, "Терминов (всего вхождений в учебнике)", sum(a1.stats_terms.values()), sum(a3.stats_terms.values()),
         sum(a3.stats_terms.values()) - sum(a1.stats_terms.values())],
        [grade, "Терминов с незнакомым ученику словом", len(nt1), len(nt3), len(nt3) - len(nt1)],
        [grade, "Знакомых слов в базе", len(known1), len(known3), len(known3) - len(known1)],
    ]
    print(f"== {grade} класс: порог 1 — {len(new1)} слов / {occ1} вхожд.; порог 3 — {len(new3)} / {occ3}; "
          f"стали выделяться {len(added)}; терминов с незнакомым словом {len(nt1)} → {len(nt3)}")

meta = morph.dictionary.meta
GUIDE = [
    ["Что сравнивается", "Две готовые версии разметки учебников: порог 1 и порог 3. Обе построены одинаково, различается только порог."],
    ["Что такое порог", "Сколько раз слово должно встретиться в прошлых классах (школьные учебники всех предметов и учебники биологии), чтобы считаться знакомым ученику. Порог 1: хотя бы раз. Порог 3: три раза и больше. Чем выше порог, тем меньше слов считаются знакомыми и тем больше слов выделено как новые."],
    ["Сводка", "Числа по каждому классу: сколько слов и терминов выделено при пороге 1 и при пороге 3. Число терминов от порога не зависит: меняется только пометка «термин с незнакомым словом». Разные падежные формы одного слова (например, «АТФ-синтетазы» и «АТФ-синтетазой») считаются одним словом."],
    ["Стали выделяться 10 кл / 11 кл", "Слова, которые при пороге 1 не выделяются (в прошлых классах встретились 1–2 раза), а при пороге 3 выделяются как новые. Для каждого слова указано, на каких страницах учебника оно стоит и сколько раз и в каких классах встречалось раньше."],
    ["Термины с незнакомым словом", "Термины из словаря терминов, в которых при пороге 3 появилось незнакомое ученику слово (при пороге 1 все слова термина считались знакомыми). Сам термин в учебнике выделен жёлтым в обеих версиях; меняется только пометка в отчётах."],
    ["Страницы", "Страницы печатного учебника (по PDF издательства)."],
    ["Упоминаний в прошлых классах", "Сколько раз слово встретилось во всех предыдущих классах. Для каждого класса берётся большее из двух чисел: по всем школьным учебникам этого класса и по учебнику биологии, чтобы одно упоминание не считалось дважды."],
    ["Словарь OpenCorpora (pymorphy3)", f"Морфологический словарь русского языка OpenCorpora (версия {meta.get('source_version')}, {meta.get('source_lexemes_count'):,} слов), который использует библиотека pymorphy3 для приведения слов к начальной форме. Если слова в нём нет, это обычно редкий научный термин или слово, испорченное при конвертации PDF в docx.".replace(",", " ")],
    ["Частота в языке (словарь Шарова)", "Частотный словарь современного русского языка О. Н. Ляшевской и С. А. Шарова (по данным Национального корпуса русского языка): сколько раз слово встречается на миллион слов. Пустая ячейка: слова в этом словаре нет."],
]

with pd.ExcelWriter(OUT, engine="openpyxl") as w:
    pd.DataFrame(GUIDE, columns=["Вкладка / понятие", "Что это"]).to_excel(w, sheet_name="Что на вкладках", index=False)
    pd.DataFrame(summary, columns=["Класс", "Показатель", "Порог 1", "Порог 3", "Разница"]).to_excel(w, sheet_name="Сводка", index=False)
    for g in sheets_words:
        sheets_words[g].to_excel(w, sheet_name=f"Стали выделяться {g} кл", index=False)
    pd.DataFrame(term_rows, columns=["Класс", "Термин", "Незнакомое слово в термине", "Упоминаний этого слова в прошлых классах",
                                     "Вхождений термина в учебнике", "Страницы"]).to_excel(w, sheet_name="Термины с незнакомым словом", index=False)

wb = load_workbook(OUT)
widths = {
    "Что на вкладках": (34, 130),
    "Сводка": (8, 46, 12, 12, 12),
    "Термины с незнакомым словом": (8, 36, 30, 22, 22, 30),
}
for name in wb.sheetnames:
    ws = wb[name]
    ws.freeze_panes = "A2"
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, width in enumerate(widths.get(name, (26, 14, 22, 16, 34, 62, 18, 18, 110)), 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    if name == "Что на вкладках":
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
wb.save(OUT)
print("saved", OUT, "| листы:", wb.sheetnames)
