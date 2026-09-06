# -*- coding: utf-8 -*-
"""Поиск мусора в результатах — отдельно по .docx и по PDF.

Автоматически ловятся только те виды мусора, у которых есть формальный
признак:

* слово разорвано переносом при извлечении PDF («представ лен», «поля - ки»);
* остатки титула и выходных данных;
* совпадение с фамилией или топонимом: словоформа с заглавной буквы в
  середине предложения, у которой лемма единицы получается только обычным
  разбором («Соловьёв» -> «соловей»);
* совпадение по варианту, а не по полному названию, — там ошибиться легче.

Смысловой мусор (омонимия: «Арктика» как регион вместо ледокола) формального
признака не имеет, поэтому в конце печатается случайная выборка для просмотра
глазами.

Запуск:  python tools/qa_noise.py [размер выборки]
"""

from __future__ import annotations

import glob
import random
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from src.lemmatizers import word_lemmas, word_proper_lemmas

#: слово, разорванное переносом строки при извлечении PDF: «поля - ки»,
#: «пред ‑ ставлен». Проверяется только для PDF — в .docx таких разрывов нет.
BROKEN = re.compile(r"[а-яё] [‑-] [а-яё]|(?<= )[а-яё]{1,3} [а-яё]{2,4} (?=[а-яё]{1,3} )")

#: титул, колофон, выходные данные
IMPRINT = re.compile(
    r"УДК|ББК|ISBN|Учебное издание|Минпросвещения|ПРОСВЕЩЕНИЕ\s*•|Тираж|Гарнитура"
)


def read(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = (
        [r for r in wb[sheet].iter_rows(min_row=2, values_only=True) if r and r[0]]
        if sheet in wb.sheetnames
        else []
    )
    wb.close()
    return rows


def looks_like_other_name(variant: str, found: str) -> bool:
    """Совпало ли слово с чужим именем собственным (фамилией, топонимом).

    Сверяется ВАРИАНТ поиска, а не полное наименование: «Атомный ледокол
    «Ленин»» находится по варианту «Ленин», и сравнивать надо с ним.

    Позицию в предложении отсюда не видно, поэтому заглавная в начале фразы
    тоже попадает в список — это верхняя оценка, а не готовый список ошибок.
    """
    if not found or not found[:1].isupper() or " " in found:
        return False
    proper = word_proper_lemmas(found)
    if not proper:
        return False
    own = word_lemmas(variant.strip("«»\"'()"))
    return not (own & proper)


def main(argv: list[str]) -> int:
    sample_size = int(argv[0]) if argv else 40

    book_files = sorted(config.OUT_PER_BOOK.rglob("*_результат.xlsx"))
    print(f"Файлов по книгам: {len(book_files)}")

    rows = []
    for path in book_files:
        for r in read(path, "Контексты"):
            rows.append((path.stem, r))

    docx_rows = [(s, r) for s, r in rows if r[4] not in ("PDF-добавка", "ручная проверка")]
    pdf_rows = [(s, r) for s, r in rows if r[4] == "PDF-добавка"]

    print(f"строк «Контексты»: всего {len(rows)}, из .docx {len(docx_rows)}, "
          f"из PDF-добавки {len(pdf_rows)}")

    for title, subset in (("ТЕКСТ .DOCX", docx_rows), ("ДОБАВКА ИЗ PDF", pdf_rows)):
        print()
        print("=" * 78)
        print(title)
        if not subset:
            print("  пусто")
            continue

        is_pdf = subset is pdf_rows
        broken = (
            [(s, r) for s, r in subset if BROKEN.search(r[6] or "")] if is_pdf else []
        )
        imprint = [(s, r) for s, r in subset if IMPRINT.search(r[6] or "")]
        foreign = [
            (s, r) for s, r in subset if looks_like_other_name(r[3] or "", r[5] or "")
        ]
        by_variant = [(s, r) for s, r in subset if r[4] not in ("full", "PDF-добавка")]

        print(f"  разорванные переносом слова рядом: "
              f"{len(broken) if is_pdf else 'не проверяется для .docx'}")
        for s, r in broken[:6]:
            print(f"     {r[0][:24]:26} | {(r[5] or '')[:18]:20} | {(r[6] or '')[:70]}")
        print(f"  остатки титула и выходных данных:  {len(imprint)}")
        for s, r in imprint[:6]:
            print(f"     {r[0][:24]:26} | {(r[5] or '')[:18]:20} | {(r[6] or '')[:70]}")
        print(f"  похоже на чужое имя собственное:   {len(foreign)}")
        counter = Counter((r[0], r[3], r[5]) for s, r in foreign)
        for (unit, variant, found), n in counter.most_common(12):
            print(f"     {n:4}  {unit[:24]:26} (вариант {variant[:14]:16}) <- {found}")
        print(f"  найдено не по полному названию, а по варианту: {len(by_variant)}")
        kinds = Counter(r[4] for s, r in by_variant)
        print(f"     {dict(kinds)}")

        print()
        print(f"  --- случайная выборка {sample_size} строк ---")
        random.seed(3)
        for s, r in random.sample(subset, min(sample_size, len(subset))):
            print(f"  [{r[4][:12]:14}] {r[0][:22]:24} | {(r[5] or '')[:20]:22} "
                  f"| {(r[6] or '')[:82]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
