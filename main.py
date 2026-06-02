"""
Точка входа: запуск полного пайплайна анализа терминов.

Что делает:
    1. Читает словарь терминов из xlsx
    2. Нормализует каждый термин -> паттерн лемм
    3. Для каждого класса (5-11):
        a. Читает текст учебника из docx
        b. Считает вхождения всех паттернов
        c. Вычисляет нормированную частоту
        d. Сохраняет контексты найденных терминов в CSV
        e. Создаёт копию учебника с подсвеченными терминами
    4. Сохраняет сводную таблицу динамики в xlsx

Запуск:
    python main.py

Для отладки отдельного термина:
    python main.py --debug-term "яйцеклетка"
"""

import argparse
import os
from collections import defaultdict
from pathlib import Path

import pandas as pd
from docx import Document

import config
from config import (
    DICT_PATH,
    GRADE_DOCX,
    GRADES,
    OUT_XLSX,
    OUT_BOOKS_DIR,
    LOG_DIR,
    FALLBACK_TOTALS,
    STOPWORDS_PATH,
    REPLACEMENTS_PATH,
    ALLOW_PUNCT_BETWEEN,
    DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
    CONTEXT_WINDOW,
    MAX_CONTEXTS_PER_TERM,
    DEBUG_ALL_TERMS,
    DEBUG_TERMS,
    NORM_PER,
)
from src.preprocessing import TextPreprocessorWithoutStopwords
from src.analyzer import (
    tokens_with_punct,
    normalize_term_to_pattern,
    count_patterns_strict,
    write_context_logs,
    debug_term_in_text,
)
from src.highlighter import annotate_docx


# =============================================================================
# Вспомогательные функции
# =============================================================================

def read_dict_terms(dict_path: str | Path) -> pd.DataFrame:
    """
    Читает словарь терминов из xlsx-файла.

    Ожидаемый формат: один столбец без заголовка,
    каждая строка — один термин.

    Args:
        dict_path: путь к файлу словаря

    Returns:
        DataFrame со столбцом 'Термин', дубликаты удалены
    """
    df = pd.read_excel(str(dict_path), header=None, engine="openpyxl")
    df = df.iloc[:, [0]].copy()
    df.columns = ['Термин']
    df['Термин'] = df['Термин'].astype(str).str.strip()
    df = df[df['Термин'].ne('') & df['Термин'].ne('nan')]
    df = df.drop_duplicates(subset=['Термин']).reset_index(drop=True)
    return df


def docx_to_text(docx_path: str | Path) -> str:
    """
    Извлекает весь текст из docx-файла.

    Обрабатывает абзацы и ячейки таблиц.
    Абзацы разделяются переносом строки.

    Args:
        docx_path: путь к docx-файлу

    Returns:
        Полный текст документа одной строкой
    """
    doc = Document(str(docx_path))
    texts = []

    for p in doc.paragraphs:
        if p.text and p.text.strip():
            texts.append(p.text)

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text for c in row.cells if c.text and c.text.strip()]
            if cells:
                texts.append(" ".join(cells))

    return "\n".join(texts)


def get_total_words(
    path: str | Path,
    tokens: list[dict],
    fallback_totals: dict[int, int],
    grade: int,
) -> int:
    """
    Определяет общее количество слов в учебнике.

    Приоритет:
        1. Число в конце имени файла (Биология_5_35834.docx -> 35834)
        2. Значение из fallback_totals[grade]
        3. Подсчёт по токенам (количество токенов с леммой)

    Args:
        path: путь к файлу
        tokens: токенизированный текст
        fallback_totals: словарь {класс: количество слов}
        grade: номер класса

    Returns:
        Количество слов (> 0)
    """
    import re
    m = re.search(r'_(\d+)(?:\.\w+)?$', str(path))
    if m:
        return int(m.group(1))

    if grade in fallback_totals and fallback_totals[grade]:
        return fallback_totals[grade]

    # Автоподсчёт по токенам
    return max(1, sum(1 for t in tokens if t.get('lemma') is not None))


# =============================================================================
# Основной пайплайн
# =============================================================================

def build_everything(
    dict_path: str | Path = DICT_PATH,
    grade_docx: dict = GRADE_DOCX,
    out_xlsx: str | Path = OUT_XLSX,
    out_books_dir: str | Path = OUT_BOOKS_DIR,
    log_dir: str | Path = LOG_DIR,
    fallback_totals: dict = FALLBACK_TOTALS,
    grades: tuple = GRADES,
    norm_per: int = NORM_PER,
) -> pd.DataFrame:
    """
    Полный пайплайн: от словаря терминов до таблицы динамики и учебников.

    Args:
        dict_path: путь к xlsx со словарём терминов
        grade_docx: словарь {класс: путь_к_docx}
        out_xlsx: путь для сохранения итоговой таблицы
        out_books_dir: директория для аннотированных учебников
        log_dir: директория для CSV с контекстами
        fallback_totals: словарь {класс: кол-во слов} для нормировки
        grades: кортеж обрабатываемых классов
        norm_per: знаменатель нормировки (стандарт: 1_000_000)

    Returns:
        DataFrame с итоговой таблицей динамики
    """
    # ------------------------------------------------------------------
    # Инициализация
    # ------------------------------------------------------------------
    print("Инициализация препроцессора...")
    preprocessor = TextPreprocessorWithoutStopwords(
        stopwords_path=STOPWORDS_PATH,
        replacements_path=REPLACEMENTS_PATH,
    )

    # ------------------------------------------------------------------
    # Загрузка и нормализация словаря терминов
    # ------------------------------------------------------------------
    print(f"Читаю словарь терминов: {dict_path}")
    terms_df = read_dict_terms(dict_path)

    terms_df["__lemmas"] = None
    terms_df["__pattern"] = None

    patterns_by_len: dict[int, list[str]] = defaultdict(list)
    pattern_to_terms: dict[str, list[str]] = defaultdict(list)
    empty_terms = 0

    print("Нормализую термины...")
    for i, term in terms_df["Термин"].items():
        lemmas, pattern = normalize_term_to_pattern(term, preprocessor)
        terms_df.at[i, "__lemmas"] = lemmas
        terms_df.at[i, "__pattern"] = pattern

        if pattern:
            patterns_by_len[len(lemmas)].append(pattern)
            pattern_to_terms[pattern].append(term)
        else:
            empty_terms += 1

    if empty_terms:
        print(
            f"[Предупреждение] {empty_terms} термин(ов) после нормализации "
            f"дали пустой паттерн — они будут иметь нулевые частоты."
        )

    # ------------------------------------------------------------------
    # Определяем паттерны для логирования контекстов
    # ------------------------------------------------------------------
    if DEBUG_ALL_TERMS:
        debug_patterns = {p for pats in patterns_by_len.values() for p in pats}
    else:
        debug_patterns = set()
        for t in DEBUG_TERMS:
            _, p = normalize_term_to_pattern(t, preprocessor)
            if p:
                debug_patterns.add(p)

    # ------------------------------------------------------------------
    # Подготовка итоговой таблицы
    # ------------------------------------------------------------------
    kv_cols = [f"КВ {g} класс" for g in grades]   # количество вхождений
    nch_cols = [f"НЧ {g} класс" for g in grades]  # нормированная частота

    for col in kv_cols + nch_cols:
        terms_df[col] = 0.0

    # ------------------------------------------------------------------
    # Создаём директории
    # ------------------------------------------------------------------
    os.makedirs(Path(out_xlsx).parent, exist_ok=True)
    os.makedirs(out_books_dir, exist_ok=True)
    if debug_patterns:
        os.makedirs(log_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # Обработка учебников по классам
    # ------------------------------------------------------------------
    for grade in grades:
        path = grade_docx.get(grade)

        if not path or not Path(path).exists():
            raise FileNotFoundError(
                f"Учебник для {grade} класса не найден: {path}\n"
                f"Проверьте TEXTBOOKS_DIR в config.py или переменную окружения "
                f"TERM_ANALYZER_TEXTBOOKS."
            )

        print(f"\n[{grade} класс] Обрабатываю: {path}")

        # Извлечение текста и токенизация
        raw_text = docx_to_text(path)
        tokens = tokens_with_punct(raw_text, preprocessor)

        total_words = get_total_words(path, tokens, fallback_totals, grade)
        print(f"[{grade} класс] Всего слов для нормировки: {total_words:,}")

        # Подсчёт вхождений
        counts_map, collected = count_patterns_strict(
            tokens,
            patterns_by_len,
            allow_punct_between=ALLOW_PUNCT_BETWEEN,
            disallow_hyphen_adjacent_for_unigrams=DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
            collect_patterns=debug_patterns,
            context_window=CONTEXT_WINDOW,
            max_collect_per_pattern=MAX_CONTEXTS_PER_TERM,
        )

        # Логирование контекстов
        if debug_patterns:
            write_context_logs(grade, collected, pattern_to_terms, log_dir, path)

        # Заполнение таблицы
        kv_col = f"КВ {grade} класс"
        nch_col = f"НЧ {grade} класс"
        denom = max(total_words, 1)

        for i, pattern in terms_df["__pattern"].items():
            count = counts_map.get(pattern, 0) if pattern else 0
            terms_df.at[i, kv_col] = int(count)
            terms_df.at[i, nch_col] = (count / denom) * norm_per

        # Аннотированный учебник
        base_name = Path(path).stem
        out_path = Path(out_books_dir) / f"{base_name} — термины.docx"
        annotate_docx(
            path,
            out_path,
            patterns_by_len,
            preprocessor,
            allow_punct_between=ALLOW_PUNCT_BETWEEN,
            disallow_hyphen_adjacent_for_unigrams=DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
        )

    # ------------------------------------------------------------------
    # Финализация типов и сохранение Excel
    # ------------------------------------------------------------------
    for col in kv_cols:
        terms_df[col] = pd.to_numeric(terms_df[col], errors='coerce').fillna(0).astype(int)
    for col in nch_cols:
        terms_df[col] = pd.to_numeric(terms_df[col], errors='coerce').fillna(0.0)

    order_cols = ["Термин"] + kv_cols + nch_cols
    result = terms_df[order_cols].copy()

    sheet_name = f"Динамика {min(grades)}-{max(grades)}"
    with pd.ExcelWriter(str(out_xlsx), engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name=sheet_name)

    print(f"\n✓ Таблица динамики:   {out_xlsx}")
    print(f"✓ Учебники:           {out_books_dir}")
    if debug_patterns:
        print(f"✓ Логи контекстов:    {log_dir}")

    return result


# =============================================================================
# CLI
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Анализ частотности биологических терминов в учебниках 5–11 классов."
    )
    parser.add_argument(
        "--debug-term",
        type=str,
        default=None,
        metavar="ТЕРМИН",
        help=(
            "Режим отладки: проверить один термин на тестовом тексте. "
            "Полный пайплайн не запускается. "
            "Пример: --debug-term \"яйцеклетка\""
        ),
    )
    parser.add_argument(
        "--debug-text",
        type=str,
        default=(
            "Рис. 4. Спорообразование у растений: женский половой орган "
            "архегоний с яйцеклеткой, дробление зиготы, молодой спорофит."
        ),
        metavar="ТЕКСТ",
        help="Текст для отладки (используется вместе с --debug-term).",
    )
    parser.add_argument(
        "--config",
        action="store_true",
        help="Вывести текущую конфигурацию и завершить работу.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.config:
        config.print_config_summary()
        return

    if args.debug_term:
        # Режим отладки одного термина
        preprocessor = TextPreprocessorWithoutStopwords(
            stopwords_path=STOPWORDS_PATH,
            replacements_path=REPLACEMENTS_PATH,
        )
        debug_term_in_text(
            term_str=args.debug_term,
            text=args.debug_text,
            preprocessor=preprocessor,
            allow_punct_between=ALLOW_PUNCT_BETWEEN,
            disallow_hyphen_adjacent_for_unigrams=DISALLOW_HYPHEN_ADJACENT_FOR_UNIGRAMS,
            window=CONTEXT_WINDOW,
        )
        return

    # Проверка конфигурации перед запуском
    issues = config.validate_config()
    if issues:
        print("Обнаружены проблемы с конфигурацией:")
        for w in issues:
            print(f"  ⚠  {w}")
        print(
            "\nПроверьте пути в config.py или задайте переменные окружения:\n"
            "  TERM_ANALYZER_DICT       — путь к словарю терминов\n"
            "  TERM_ANALYZER_TEXTBOOKS  — директория с учебниками\n"
            "  TERM_ANALYZER_OUTPUT     — директория для результатов\n"
        )
        return

    # Полный пайплайн
    build_everything()


if __name__ == "__main__":
    main()