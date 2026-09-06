# -*- coding: utf-8 -*-
"""Поиск единиц культурного кода в учебниках истории.

Примеры:

    python main.py                                   # оба бэкенда, все учебники
    python main.py --lemmatizer pymorphy             # только pymorphy3
    python main.py --only 05_01 --no-highlight       # быстрый прогон одного учебника
    python main.py --debug-unit "Крейсер «Аврора»"   # как разбирается наименование
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import config
from src import analyzer, pdfsource, pipeline, report
from src.lemmatizers import get_backend
from src.units import build_variants, load_units


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Поиск единиц культурного кода")
    parser.add_argument(
        "--lemmatizer",
        choices=["pymorphy", "natasha", "both"],
        default="both",
        help="какой вариант лемматизации использовать (по умолчанию оба со сравнением)",
    )
    parser.add_argument(
        "--only",
        action="append",
        default=None,
        help="обработать только учебники, чьё имя содержит подстроку (можно повторять)",
    )
    parser.add_argument(
        "--no-highlight", action="store_true", help="не создавать docx с подсветкой"
    )
    parser.add_argument(
        "--primary",
        choices=["pymorphy", "natasha"],
        default=config.PRIMARY_BACKEND,
        help="бэкенд, по которому собираются основные отчёты",
    )
    parser.add_argument(
        "--no-pdf",
        action="store_true",
        help="не разбирать PDF (по умолчанию считается лист «Только в PDF»)",
    )
    parser.add_argument(
        "--debug-unit", default=None, help="показать варианты и паттерны наименования"
    )
    return parser.parse_args(argv)


def log(message: str) -> None:
    print(message, flush=True)


def _partial(path: Path) -> Path:
    """Имя файла для частичного прогона."""
    return path.with_name(f"{path.stem} (частичный){path.suffix}")


def debug_unit(name: str, units) -> None:
    """Печатает варианты поиска и лемма-паттерны для одного наименования."""
    for backend_name in ("pymorphy", "natasha"):
        backend = get_backend(backend_name)
        build_variants(units, backend)
        matched = [u for u in units if u.name == name]
        if not matched:
            lowered = name.lower()
            matched = [u for u in units if lowered in u.name.lower()]
        log(f"\n=== {backend_name} ===")
        for unit in matched:
            log(f"{unit.name}  [{unit.group_name}]")
            for variant in unit.variants:
                log(f"    {variant.kind:12} {variant.surface!r} -> {' '.join(variant.pattern)}")
        if not matched:
            log(f"  наименование {name!r} в списке не найдено")


def run_backend(backend_name: str, units, books, is_primary: bool, args):
    """Прогоняет корпус одним бэкендом, возвращает (результаты, леммы, частоты)."""
    backend = get_backend(backend_name)
    log(f"\n=== Лемматизация: {backend_name} ===")

    t0 = time.perf_counter()
    build_variants(units, backend)
    index = analyzer.build_pattern_index(units)
    first_lemmas = analyzer.first_lemma_set(index)
    variants = sum(len(u.variants) for u in units)
    log(
        f"  единиц: {len(units)}, вариантов поиска: {variants}, "
        f"первых лемм: {len(first_lemmas)}, макс. длина: "
        f"{analyzer.max_pattern_length(index)} ({time.perf_counter() - t0:.1f} с)"
    )

    lemmas: dict[str, str] = {}
    freq: dict[str, int] = {}
    results = []

    for book in books:
        highlight_out = None
        if is_primary and not args.no_highlight:
            highlight_out = config.OUT_HIGHLIGHT / book.parent.name / book.name

        result = pipeline.process_book(
            book, units, index, first_lemmas, backend,
            highlight_out=highlight_out, lemma_sink=lemmas, form_freq=freq,
        )
        results.append(result)

        note = ""
        if result.highlight_stats:
            hs = result.highlight_stats
            note = f", подсвечено {hs.spans_applied} в {hs.paragraphs_touched} абз."
            if hs.spans_skipped_complex_run:
                note += f" (пропущено {hs.spans_skipped_complex_run} в сложных run)"
        log(
            f"  {result.rel_path}: {result.total_hits} вхождений, "
            f"{result.unique_units} единиц, {result.elapsed:.1f} с{note}"
        )

    # PDF разбирается до записи книжных отчётов: его вхождения складываются
    # с основными, поэтому книжные файлы должны их уже содержать.
    pdf_results = []
    if is_primary and not args.no_pdf:
        pairs = pdfsource.pair_pdfs_with_docx(books)
        if pairs:
            log(f"  Разбираю PDF ({len(pairs)}) — то, чего нет в .docx…")
        by_pair = {res.pair_key: res for res in results if res.pair_key}
        for pdf, docx in pairs:
            res = pipeline.process_pdf_extra(
                pdf, docx, units, index, first_lemmas, backend
            )
            pdf_results.append(res)
            merged = ""
            if config.COUNT_PDF_EXTRA:
                book_result = by_pair.get(res.pair_key)
                if book_result is not None:
                    pipeline.merge_pdf_extra(book_result, res, units)
                    merged = " — сложено с основными"
            log(
                f"    {res.rel_path}: {res.paragraphs} фрагментов вне .docx, "
                f"{res.total_hits} вхождений, {res.unique_units} единиц, "
                f"{res.elapsed:.1f} с{merged}"
            )

    if is_primary:
        for book, result in zip(books, results):
            out = config.OUT_PER_BOOK / book.parent.name / f"{book.stem}_результат.xlsx"
            report.write_book_report(result, units, out)

    return results, lemmas, freq, pdf_results


def main(argv=None) -> int:
    args = parse_args(argv)

    log("Читаю список единиц культурного кода…")
    units = load_units()
    groups = sorted({u.group_sheet for u in units})
    log(f"  пар «наименование × группа»: {len(units)}, групп: {len(groups)}")

    if args.debug_unit:
        debug_unit(args.debug_unit, units)
        return 0

    books = pipeline.find_books(config.CORPUS_ROOT)
    if args.only:
        books = [b for b in books if any(part in b.name for part in args.only)]
    if not books:
        log("Не найдено ни одного учебника — проверьте --only и config.CORPUS_ROOT")
        return 1
    log(f"Учебников к обработке: {len(books)}")

    config.OUT_ROOT.mkdir(parents=True, exist_ok=True)

    # Частичный прогон не должен затирать полный сводный отчёт: он собирается
    # только по обработанным книгам, и подменять им полный нельзя.
    summary_path = config.OUT_SUMMARY_XLSX
    compare_path = config.OUT_COMPARE_XLSX
    if args.only:
        summary_path = _partial(summary_path)
        compare_path = _partial(compare_path)
        log(
            f"Прогон частичный (--only): сводный отчёт пишется в "
            f"«{summary_path.name}», полный не трогаю."
        )

    backends = ["pymorphy", "natasha"] if args.lemmatizer == "both" else [args.lemmatizer]
    primary = args.primary if args.primary in backends else backends[0]

    collected = {}
    for backend_name in backends:
        collected[backend_name] = run_backend(
            backend_name, units, books, backend_name == primary, args
        )

    primary_results, _, _, pdf_results = collected[primary]
    log("\nСобираю сводный отчёт…")
    report.write_summary(
        primary_results, units, summary_path, pdf_results=pdf_results
    )
    log(f"  {summary_path}")

    if len(backends) == 2:
        log("Собираю отчёт сравнения лемматизаторов…")
        res_a, lem_a, freq_a, _ = collected["pymorphy"]
        res_b, lem_b, _, _ = collected["natasha"]
        report.write_comparison(
            res_a, res_b, units, lem_a, lem_b, freq_a, compare_path
        )
        log(f"  {compare_path}")

    total = sum(r.total_hits for r in primary_results)
    log(f"\nГотово. Всего вхождений ({primary}): {total}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
