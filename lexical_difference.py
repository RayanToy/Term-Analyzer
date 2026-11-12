"""
Лексическая разница (лексическое обогащение) учебников биологии 10 и 11 класса.

Как в DataAnalyzer.calculate_lexical_difference (Russian-Lexicon-Analyzer):
частотные списки лемм двух учебников сравниваются как множества —
леммы второго списка, которых нет в первом, и есть лексическая разница.
Здесь она считается в обе стороны и дополнена сводными числами:
длина учебников в токенах и леммах, общие леммы, доля новых лемм и их
токенов в тексте, коэффициент обогащения.

Леммы берутся тем же разбором, что и в подсветке и отчётах (Natasha +
pymorphy3, знаменательные слова; мусор распознавания исключён).

Запуск: python lexical_difference.py [--min-freq 3]
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

import config
import highlight_new_and_terms as h
import term_dictionaries as td
import textbook_report as tr
from src.analyzer import build_indexed_words

OUT = h.HIGHLIGHT_DIR / "Лексическая разница_учебники 10 и 11 класса.xlsx"


def lemma_counts(a) -> Counter:
    """Частотный список лемм учебника: лемма -> токенов (словоупотреблений)."""
    c = Counter()
    for _, _, tokens in a.blocks:
        for w in build_indexed_words(tokens):
            if w["lemma"] not in a.junk:
                c[w["lemma"]] += 1
    return c


def prior_counts(pre, grades: list[int], use_cache: bool) -> Counter:
    """Упоминания лемм во всех учебниках классов grades (за класс — максимум из двух источников, как в подсветке)."""
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


def term_words(a) -> dict[str, str]:
    """Лемма -> словари, в термины которых она входит (по найденным в учебнике терминам)."""
    sources = a.get("term_sources") or {}
    out = defaultdict(set)
    for term in a.stats_terms:
        for w in term.split():
            out[w] |= set(sources.get(term, {"dict"}))
    return {w: ", ".join(td.SOURCE_NAMES[s] for s in td.SOURCES if s in ss) for w, ss in out.items()}


def pct(x: float, y: float) -> float:
    return round(x / y * 100, 1) if y else 0.0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--min-freq", type=int, default=config.MIN_KNOWN_FREQ)
    ap.add_argument("--no-cache", action="store_true")
    args = ap.parse_args()
    use_cache = not args.no_cache

    A = {g: tr.analyze(g, args.min_freq, use_cache) for g in (10, 11)}
    C = {g: lemma_counts(A[g]) for g in A}
    tokens = {g: sum(C[g].values()) for g in C}
    tokens_all = {g: sum(1 for _, _, toks in A[g].blocks for t in toks
                         if any(ch.isalpha() for ch in t["text"])) for g in A}
    L10, L11 = set(C[10]), set(C[11])
    common, only10, only11 = L10 & L11, L10 - L11, L11 - L10
    pre = A[11].preprocessor
    before = prior_counts(pre, list(range(2, 11)), use_cache)     # всё, что было до 11 класса, включая 10-й
    before10 = prior_counts(pre, list(range(2, 10)), use_cache)   # до 10 класса
    terms = {g: term_words(A[g]) for g in A}
    new_for = {g: set(A[g].stats_new) for g in A}

    tok_only11 = sum(C[11][l] for l in only11)
    tok_only10 = sum(C[10][l] for l in only10)
    only11_unseen = [l for l in only11 if not before.get(l)]
    summary = [
        ["Метод", "Лексическая разница — как в Russian-Lexicon-Analyzer (DataAnalyzer.calculate_lexical_difference): "
                  "частотные списки лемм двух учебников сравниваются как множества; леммы, которые есть во втором "
                  "учебнике и которых нет в первом, — лексическое обогащение второго учебника."],
        ["Токен и лемма", tr.TOKENS_NOTE[:1].upper() + tr.TOKENS_NOTE[1:] + ". Считаются знаменательные слова: "
                          "предлоги, союзы, частицы, местоимения и числа не входят; мусор распознавания исключён."],
        ["Длина учебника 10 кл.: токенов (все слова текста)", tokens_all[10]],
        ["Длина учебника 10 кл.: токенов знаменательных слов", tokens[10]],
        ["Длина учебника 10 кл.: лемм", len(L10)],
        ["Длина учебника 11 кл.: токенов (все слова текста)", tokens_all[11]],
        ["Длина учебника 11 кл.: токенов знаменательных слов", tokens[11]],
        ["Длина учебника 11 кл.: лемм", len(L11)],
        ["Общих лемм (есть в обоих учебниках)", len(common)],
        ["Сходство словарей (Жаккар): общие леммы / все леммы обоих учебников, %", pct(len(common), len(L10 | L11))],
        ["Лексическое обогащение 11 кл.: лемм, которых нет в учебнике 10 кл.", len(only11)],
        ["  доля от лемм учебника 11 кл., %", pct(len(only11), len(L11))],
        ["  их токенов в тексте 11 кл.", tok_only11],
        ["  доля от токенов учебника 11 кл., %", pct(tok_only11, tokens[11])],
        ["  коэффициент обогащения: новые леммы 11 кл. / леммы 10 кл. × 100", pct(len(only11), len(L10))],
        ["  из них не встречались ни в одном школьном учебнике 2–10 классов (лемм)", len(only11_unseen)],
        ["  из них выделены в учебнике 11 кл. как новые для ученика (лемм)", len(only11 & new_for[11])],
        ["  из них входят в термины учебника 11 кл. (лемм)", sum(1 for l in only11 if l in terms[11])],
        ["Лемм учебника 10 кл., которых нет в учебнике 11 кл.", len(only10)],
        ["  доля от лемм учебника 10 кл., %", pct(len(only10), len(L10))],
        ["  их токенов в тексте 10 кл.", tok_only10],
    ]

    def rows_only(g, lemmas, prior, prior_label):
        n = tokens[g]
        rows = []
        for l in sorted(lemmas, key=lambda x: (-C[g][x], x)):
            rows.append([l, C[g][l], round(C[g][l] / n * 10000, 2), terms[g].get(l, ""),
                         "да" if l in new_for[g] else "", int(prior.get(l, 0))])
        return pd.DataFrame(rows, columns=[
            "Лемма", f"Токенов в учебнике {g} кл.", f"На 10 000 токенов учебника {g} кл.",
            f"Входит в термины учебника {g} кл. (словари)", f"Выделена в учебнике {g} кл. как новая для ученика",
            prior_label])

    df11 = rows_only(11, only11, before, "Упоминаний во всех школьных учебниках 2–10 кл. (токенов)")
    df10 = rows_only(10, only10, before10, "Упоминаний во всех школьных учебниках 2–9 кл. (токенов)")
    common_rows = [[l, C[10][l], C[11][l], round(C[10][l] / tokens[10] * 10000, 2),
                    round(C[11][l] / tokens[11] * 10000, 2)]
                   for l in sorted(common, key=lambda x: (-(C[10][x] + C[11][x]), x))]
    df_common = pd.DataFrame(common_rows, columns=[
        "Лемма", "Токенов в учебнике 10 кл.", "Токенов в учебнике 11 кл.",
        "На 10 000 токенов 10 кл.", "На 10 000 токенов 11 кл."])
    guide = [
        ["Сводка", "Длина учебников в токенах и леммах и итоговые числа лексической разницы."],
        ["Обогащение 11 (нет в 10)", "Леммы учебника 11 класса, которых нет в учебнике 10 класса, — лексическое "
                                     "обогащение 11 класса по сравнению с 10-м. Отсортированы по числу токенов."],
        ["Нет в 11 (есть в 10)", "Леммы учебника 10 класса, которые в учебнике 11 класса не встречаются."],
        ["Общие леммы", "Леммы, которые есть в обоих учебниках, с числом токенов в каждом."],
        ["На 10 000 токенов", "Частота леммы, нормированная на длину учебника: сколько раз она встретилась бы "
                              "на 10 000 токенов текста. Позволяет сравнивать учебники разной длины."],
        ["Входит в термины", "Лемма — часть термина (из ФРП, указателя учебника или словаря терминов), найденного "
                             "в этом учебнике."],
        ["Выделена как новая для ученика", f"Лемма выделена в размеченном учебнике бирюзовым: в прошлых классах "
                                           f"встречалась меньше {args.min_freq} раз (токенов)."],
        ["Упоминаний во всех школьных учебниках", "Сколько раз (токенов любой формы) лемма встретилась в школьных "
                                                   "учебниках всех предметов прошлых классов (частотные списки и "
                                                   "учебники биологии; за каждый класс берётся большее из двух чисел)."],
    ]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(str(OUT), engine="openpyxl") as w:
        pd.DataFrame(summary, columns=["Показатель", "Значение"]).to_excel(w, sheet_name="Сводка", index=False)
        df11.to_excel(w, sheet_name="Обогащение 11 (нет в 10)", index=False)
        df10.to_excel(w, sheet_name="Нет в 11 (есть в 10)", index=False)
        df_common.to_excel(w, sheet_name="Общие леммы", index=False)
        pd.DataFrame(guide, columns=["Вкладка / столбец", "Что это"]).to_excel(w, sheet_name="Пояснения", index=False)

    wb = load_workbook(OUT)
    widths = {"Сводка": [70, 60], "Пояснения": [36, 110], "Общие леммы": [26, 16, 16, 16, 16]}
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for i, c in enumerate(ws[1]):
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            col = c.column_letter
            ws.column_dimensions[col].width = widths.get(ws.title, [26, 16, 16, 30, 18, 22])[min(i, 5)] \
                if ws.title not in widths else widths[ws.title][min(i, len(widths[ws.title]) - 1)]
        if ws.title in ("Сводка", "Пояснения"):
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(OUT)
    print(f"[Лексическая разница] 10 кл.: {len(L10)} лемм / {tokens[10]} токенов; 11 кл.: {len(L11)} лемм / "
          f"{tokens[11]} токенов; общих {len(common)}; только в 11 — {len(only11)}, только в 10 — {len(only10)}")
    print(f"[Лексическая разница] Сохранено: {OUT}")


if __name__ == "__main__":
    main()
