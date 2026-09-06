# -*- coding: utf-8 -*-
"""Сборка Excel-отчётов.

Структура листов повторяет прошлую работу по фразеологизмам
(`_РЕЗУЛЬТАТ_фразеологизмы.xlsx` и `*_результат.xlsx`), плюс добавлен лист
«Спорные» — вхождения имён собственных, где заглавная буква была вынужденной.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.units import UnitEntry, _sheet_order

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)

BOOK_COLUMNS = [
    "Наименование объекта КК",
    "Тематическая группа",
    "Количество",
    "в т.ч. из PDF",
    "в т.ч. добавлено вручную",
    "из них спорных засчитано",
    "спорных отклонено",
    "отклонено: строчная буква",
]

CONTEXT_COLUMNS = [
    "Наименование объекта КК",
    "Тематическая группа",
    "#",
    "Вариант поиска",
    "Тип варианта",
    "Найдено в тексте",
    "Контекст",
    "Статус",
    "Причина",
]

DISPUTED_COLUMNS = [
    "Наименование объекта КК",
    "Тематическая группа",
    "Вариант поиска",
    "Найдено в тексте",
    "Контекст",
    "Статус",
    "Причина",
]


REVIEW_COLUMNS = [
    "Наименование объекта КК",
    "Тематическая группа",
    "Вариант поиска",
    "Что нашлось рядом в тексте",
    "Проверено вручную",
    "Комментарий",
]


# ---------------------------------------------------------------------------
# Вспомогательное
# ---------------------------------------------------------------------------


#: символы, запрещённые в xlsx (те же, что отвергает openpyxl)
ILLEGAL_XLSX = re.compile(r"[\000-\010\013\014\016-\037]")

#: длина текста в ячейке ограничена форматом
MAX_CELL_LEN = 32767


def _clean(value):
    """Готовит значение к записи в ячейку.

    Текст, извлечённый из PDF, содержит управляющие символы, на которых
    openpyxl падает с IllegalCharacterError.
    """
    if not isinstance(value, str):
        return value
    value = ILLEGAL_XLSX.sub(" ", value)
    return value[:MAX_CELL_LEN]


def save_workbook(wb, path: Path) -> Path:
    """Сохраняет книгу, не теряя результат, если файл открыт в Excel.

    Excel держит открытый файл заблокированным, и ``wb.save`` падает с
    PermissionError уже в самом конце прогона — после всех вычислений.
    В этом случае пишем рядом файл с пометкой «(новый)» и говорим об этом,
    вместо того чтобы выбросить работу.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem} (новый){path.suffix}")
        wb.save(fallback)
        print(
            f"  ВНИМАНИЕ: «{path.name}» открыт в Excel и заблокирован."
            f" Результат сохранён как «{fallback.name}»."
            f" Закройте файл и переименуйте, либо перезапустите прогон.",
            flush=True,
        )
        return fallback


def _write_sheet(ws, columns: list[str], rows, widths=None) -> None:
    ws.append(columns)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append([_clean(v) for v in row])
    ws.freeze_panes = "A2"
    widths = widths or _default_widths(columns)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def _default_widths(columns: list[str]) -> list[int]:
    out = []
    for name in columns:
        if "Контекст" in name:
            out.append(90)
        elif "Наименование" in name or "Учебник" in name or "Вариант" in name:
            out.append(38)
        elif "группа" in name or "Причина" in name or "Статус" in name:
            out.append(26)
        else:
            out.append(14)
    return out


def _unit_rows(units, counts, pdf, manual, accepted, rejected, lowercase,
               only_found=False, group=None):
    rows = []
    for idx, unit in enumerate(units):
        if group is not None and unit.group_sheet != group:
            continue
        total = counts.get(idx, 0)
        if only_found and not total:
            continue
        rows.append(
            [
                unit.name,
                unit.group_name,
                total,
                pdf.get(idx, 0),
                manual.get(idx, 0),
                accepted.get(idx, 0),
                rejected.get(idx, 0),
                lowercase.get(idx, 0),
            ]
        )
    if only_found:
        rows.sort(key=lambda r: (-r[2], r[0]))
    return rows


def _review_rows(result, units):
    """Строки листа «Возможные пропуски»."""
    rows = [
        [
            units[r["unit_idx"]].name,
            units[r["unit_idx"]].group_name,
            r["variant"],
            r["context"],
            r.get("verdict", "не проверено"),
            r.get("comment", ""),
        ]
        for r in result.review_rows
    ]
    # сначала подтверждённые пропуски, потом пограничные, потом ложные
    order = {"пропуск": 0, "пограничное": 1, "не проверено": 2, "ложное": 3}
    rows.sort(key=lambda r: (order.get(r[4], 2), r[0]))
    return rows


# ---------------------------------------------------------------------------
# Отчёт по одной книге
# ---------------------------------------------------------------------------


def write_book_report(result, units: list[UnitEntry], path: Path) -> None:
    """Excel по одному учебнику."""
    wb = Workbook()
    wb.remove(wb.active)

    _write_sheet(
        wb.create_sheet("Все единицы"),
        BOOK_COLUMNS,
        _unit_rows(units, result.counts, result.pdf, result.manual,
                   result.accepted, result.rejected, result.lowercase),
    )
    _write_sheet(
        wb.create_sheet("Найденные"),
        BOOK_COLUMNS,
        _unit_rows(
            units, result.counts, result.pdf, result.manual, result.accepted,
            result.rejected, result.lowercase, only_found=True,
        ),
    )

    context_rows = [
        [
            units[c["unit_idx"]].name,
            units[c["unit_idx"]].group_name,
            c["n"],
            c["variant"],
            c["kind"],
            c["found"],
            c["context"],
            c["status"],
            c["reason"],
        ]
        for c in result.contexts
    ]
    context_rows.sort(key=lambda r: (r[0], r[2]))
    _write_sheet(wb.create_sheet("Контексты"), CONTEXT_COLUMNS, context_rows)

    disputed_rows = [
        [
            units[d["unit_idx"]].name,
            units[d["unit_idx"]].group_name,
            d["variant"],
            d["found"],
            d["context"],
            d["status"],
            d["reason"],
        ]
        for d in result.disputed_rows
    ]
    disputed_rows.sort(key=lambda r: (r[5], r[0]))
    _write_sheet(wb.create_sheet("Спорные"), DISPUTED_COLUMNS, disputed_rows)

    _write_sheet(
        wb.create_sheet("Возможные пропуски"),
        REVIEW_COLUMNS,
        _review_rows(result, units),
    )

    for sheet in sorted({u.group_sheet for u in units}, key=_sheet_order):
        _write_sheet(
            wb.create_sheet(sheet[:31]),
            BOOK_COLUMNS,
            _unit_rows(
                units, result.counts, result.pdf, result.manual, result.accepted,
                result.rejected, result.lowercase, group=sheet,
            ),
        )

    save_workbook(wb, path)


# ---------------------------------------------------------------------------
# Сводный отчёт
# ---------------------------------------------------------------------------


def write_summary(results, units: list[UnitEntry], path: Path, pdf_results=None) -> None:
    """Сводный Excel по всем учебникам.

    ``pdf_results`` — результаты по той части PDF, которой нет в .docx;
    они выносятся на отдельный лист и в общие числа не входят.
    """
    wb = Workbook()
    wb.remove(wb.active)

    totals: dict[int, int] = defaultdict(int)
    manual: dict[int, int] = defaultdict(int)
    pdf_extra: dict[int, int] = defaultdict(int)
    accepted: dict[int, int] = defaultdict(int)
    rejected: dict[int, int] = defaultdict(int)
    lowercase: dict[int, int] = defaultdict(int)
    per_book: dict[int, dict[str, int]] = defaultdict(dict)

    for res in results:
        for idx, count in res.counts.items():
            if count:
                totals[idx] += count
                per_book[idx][res.rel_path] = count
        for idx, count in res.manual.items():
            manual[idx] += count
        for idx, count in res.pdf.items():
            pdf_extra[idx] += count
        for idx, count in res.accepted.items():
            accepted[idx] += count
        for idx, count in res.rejected.items():
            rejected[idx] += count
        for idx, count in res.lowercase.items():
            lowercase[idx] += count

    books = [res.rel_path for res in results]

    found_columns = [
        "Наименование объекта КК",
        "Тематическая группа",
        "Итого",
        "в т.ч. из PDF",
        "в т.ч. добавлено вручную",
        "В скольких учебниках",
        "из них спорных засчитано",
        "спорных отклонено",
        "отклонено: строчная буква",
        "Учебники",
    ]

    def found_rows(only_found=True, group=None):
        rows = []
        for idx, unit in enumerate(units):
            if group is not None and unit.group_sheet != group:
                continue
            total = totals.get(idx, 0)
            if only_found and not total:
                continue
            places = per_book.get(idx, {})
            rows.append(
                [
                    unit.name,
                    unit.group_name,
                    total,
                    pdf_extra.get(idx, 0),
                    manual.get(idx, 0),
                    len(places),
                    accepted.get(idx, 0),
                    rejected.get(idx, 0),
                    lowercase.get(idx, 0),
                    "\n".join(f"{book} ({n})" for book, n in sorted(places.items())),
                ]
            )
        rows.sort(key=lambda r: (-r[2], r[0]))
        return rows

    _write_sheet(wb.create_sheet("Найденные"), found_columns, found_rows())

    pair_rows = [
        [units[idx].name, units[idx].group_name, book, count]
        for idx, places in sorted(per_book.items())
        for book, count in sorted(places.items())
    ]
    pair_rows.sort(key=lambda r: (r[0], r[2]))
    _write_sheet(
        wb.create_sheet("Единица × Учебник"),
        ["Наименование объекта КК", "Тематическая группа", "Учебник", "Количество"],
        pair_rows,
    )

    matrix_columns = ["Наименование объекта КК", "Тематическая группа", "Итого"] + books
    matrix_rows = []
    for idx, unit in enumerate(units):
        total = totals.get(idx, 0)
        if not total:
            continue
        places = per_book.get(idx, {})
        matrix_rows.append(
            [unit.name, unit.group_name, total] + [places.get(b, 0) for b in books]
        )
    matrix_rows.sort(key=lambda r: (-r[2], r[0]))
    _write_sheet(
        wb.create_sheet("Матрица"),
        matrix_columns,
        matrix_rows,
        widths=[38, 26, 10] + [22] * len(books),
    )

    _write_sheet(wb.create_sheet("Все единицы"), found_columns, found_rows(only_found=False))

    _write_sheet(
        wb.create_sheet("Сводка по учебникам"),
        [
            "Учебник",
            "Вхождений (строк отчёта)",
            "Вхождений (мест в тексте)",
            "в т.ч. из PDF",
            "в т.ч. добавлено вручную",
            "Уник. единиц",
            "Абзацев",
            "Словарных токенов",
            "Время, с",
        ],
        [
            [
                res.rel_path,
                res.total_hits,
                res.distinct_hits,
                sum(res.pdf.values()),
                sum(res.manual.values()),
                res.unique_units,
                res.paragraphs,
                res.total_words,
                round(res.elapsed, 1),
            ]
            for res in results
        ],
        widths=[60, 22, 22, 18, 22, 14, 12, 20, 12],
    )

    disputed_rows = [
        [
            res.rel_path,
            units[d["unit_idx"]].name,
            units[d["unit_idx"]].group_name,
            d["variant"],
            d["found"],
            d["context"],
            d["status"],
            d["reason"],
        ]
        for res in results
        for d in res.disputed_rows
    ]
    disputed_rows.sort(key=lambda r: (r[6], r[1], r[0]))
    _write_sheet(
        wb.create_sheet("Спорные"),
        ["Учебник"] + DISPUTED_COLUMNS,
        disputed_rows,
    )

    review_rows = [
        [
            res.rel_path,
            units[r["unit_idx"]].name,
            units[r["unit_idx"]].group_name,
            r["variant"],
            r["context"],
            r.get("verdict", "не проверено"),
            r.get("comment", ""),
        ]
        for res in results
        for r in res.review_rows
    ]
    order = {"пропуск": 0, "пограничное": 1, "не проверено": 2, "ложное": 3}
    review_rows.sort(key=lambda r: (order.get(r[5], 2), r[1], r[0]))
    _write_sheet(
        wb.create_sheet("Возможные пропуски"),
        ["Учебник"] + REVIEW_COLUMNS,
        review_rows,
        widths=[60, 38, 26, 34, 100, 20, 60],
    )

    if pdf_results:
        _write_pdf_sheet(wb, pdf_results, units, results)

    for sheet in sorted({u.group_sheet for u in units}, key=_sheet_order):
        _write_sheet(
            wb.create_sheet(sheet[:31]), found_columns, found_rows(only_found=False, group=sheet)
        )

    save_workbook(wb, path)


def _write_pdf_sheet(wb, pdf_results, units: list[UnitEntry], docx_results) -> None:
    """Лист «Только в PDF»: единицы из подписей к иллюстрациям и колонтитулов.

    Колонка «Есть ли в .docx» показывает, сколько раз та же единица нашлась
    в основном тексте того же учебника: ноль означает, что без PDF единица
    была бы упущена совсем.

    Текст извлечён из PDF, порядок слов местами нарушен, поэтому лист носит
    справочный характер и в основные числа не входит.
    """
    docx_by_pair = {res.pair_key: res for res in docx_results if res.pair_key}

    rows = []
    for res in pdf_results:
        docx_res = docx_by_pair.get(res.pair_key)
        samples = {}
        for c in res.contexts:
            samples.setdefault(c["unit_idx"], c)
        for idx, count in sorted(res.counts.items()):
            if not count:
                continue
            sample = samples.get(idx)
            # Вклад PDF уже сложен с основными числами, поэтому «сколько раз
            # единица встречается в самом .docx» считаем как разность.
            in_docx = (
                docx_res.counts.get(idx, 0) - docx_res.pdf.get(idx, 0)
                if docx_res
                else None
            )
            rows.append(
                [
                    res.rel_path,
                    units[idx].name,
                    units[idx].group_name,
                    count,
                    in_docx,
                    "да" if in_docx else "НЕТ — только в PDF",
                    sample["found"] if sample else "",
                    sample["context"] if sample else "",
                ]
            )
    # Сначала то, чего в .docx нет вовсе, — ради этого лист и делается.
    rows.sort(key=lambda r: (r[4] or 0, -r[3], r[1]))
    _write_sheet(
        wb.create_sheet("Только в PDF"),
        [
            "PDF-файл",
            "Наименование объекта КК",
            "Тематическая группа",
            "Вхождений в PDF-добавке",
            "Вхождений в .docx",
            "Есть ли в .docx",
            "Пример найденного",
            "Пример контекста",
        ],
        rows,
        widths=[60, 38, 26, 20, 18, 22, 28, 90],
    )


# ---------------------------------------------------------------------------
# Сравнение бэкендов
# ---------------------------------------------------------------------------


def write_comparison(results_a, results_b, units, lemmas_a, lemmas_b, freq, path: Path) -> None:
    """Отчёт сравнения двух вариантов лемматизации.

    ``results_a`` / ``results_b`` — списки ``BookResult`` по одним и тем же
    учебникам в одном порядке.
    """
    name_a = results_a[0].backend if results_a else "A"
    name_b = results_b[0].backend if results_b else "B"

    wb = Workbook()
    wb.remove(wb.active)

    by_book_b = {res.rel_path: res for res in results_b}

    summary_rows = []
    for res_a in results_a:
        res_b = by_book_b.get(res_a.rel_path)
        if res_b is None:
            continue
        keys_a = set(res_a.match_keys)
        keys_b = set(res_b.match_keys)
        summary_rows.append(
            [
                res_a.rel_path,
                res_a.total_words,
                res_b.total_words,
                res_a.total_hits,
                res_b.total_hits,
                res_a.total_hits - res_b.total_hits,
                len(keys_a - keys_b),
                len(keys_b - keys_a),
                round(res_a.elapsed, 1),
                round(res_b.elapsed, 1),
            ]
        )
    _write_sheet(
        wb.create_sheet("Сводка"),
        [
            "Учебник",
            f"Словарных токенов, {name_a}",
            f"Словарных токенов, {name_b}",
            f"Вхождений, {name_a}",
            f"Вхождений, {name_b}",
            "Δ вхождений",
            f"Только {name_a}",
            f"Только {name_b}",
            f"Время {name_a}, с",
            f"Время {name_b}, с",
        ],
        summary_rows,
        widths=[60] + [18] * 9,
    )

    diff_rows = []
    for res_a in results_a:
        res_b = by_book_b.get(res_a.rel_path)
        if res_b is None:
            continue
        for idx in set(res_a.counts) | set(res_b.counts):
            a = res_a.counts.get(idx, 0)
            b = res_b.counts.get(idx, 0)
            if a != b:
                diff_rows.append(
                    [units[idx].name, units[idx].group_name, res_a.rel_path, a, b, a - b]
                )
    diff_rows.sort(key=lambda r: (-abs(r[5]), r[0]))
    _write_sheet(
        wb.create_sheet("Различия по единицам"),
        [
            "Наименование объекта КК",
            "Тематическая группа",
            "Учебник",
            f"Кол-во, {name_a}",
            f"Кол-во, {name_b}",
            "Δ",
        ],
        diff_rows,
        widths=[38, 26, 60, 16, 16, 10],
    )

    for label, first, second in (
        (f"Только {name_a}", results_a, by_book_b),
        (f"Только {name_b}", results_b, {res.rel_path: res for res in results_a}),
    ):
        rows = []
        for res in first:
            other = second.get(res.rel_path)
            if other is None:
                continue
            other_keys = set(other.match_keys)
            lookup = {
                (c["unit_idx"], c["found"]): c
                for c in res.contexts
            }
            for key, status in res.match_keys.items():
                if key in other_keys:
                    continue
                unit_idx = key[0]
                sample = next(
                    (c for (u, _f), c in lookup.items() if u == unit_idx), None
                )
                rows.append(
                    [
                        res.rel_path,
                        units[unit_idx].name,
                        units[unit_idx].group_name,
                        status,
                        sample["found"] if sample else "",
                        sample["context"] if sample else "",
                    ]
                )
        rows.sort(key=lambda r: (r[1], r[0]))
        _write_sheet(
            wb.create_sheet(label[:31]),
            ["Учебник", "Наименование объекта КК", "Тематическая группа", "Статус",
             "Пример найденного", "Пример контекста"],
            rows[:50000],
            widths=[60, 38, 26, 24, 28, 90],
        )

    lemma_rows = []
    for form, lemma_a in lemmas_a.items():
        lemma_b = lemmas_b.get(form)
        if lemma_b is not None and lemma_b != lemma_a:
            lemma_rows.append([form, lemma_a, lemma_b, freq.get(form, 0)])
    lemma_rows.sort(key=lambda r: -r[3])
    _write_sheet(
        wb.create_sheet("Спорные леммы"),
        ["Словоформа", f"Лемма, {name_a}", f"Лемма, {name_b}", "Частота в корпусе"],
        lemma_rows[:50000],
        widths=[26, 26, 26, 20],
    )

    save_workbook(wb, path)
