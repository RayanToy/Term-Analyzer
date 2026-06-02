"""
Анализ текста: токенизация с лемматизацией, поиск паттернов терминов,
подсчёт вхождений, сбор контекстов для логирования.

Основной поток данных:
    raw text
        -> tokens_with_punct()          # токены с леммами и позициями
        -> build_indexed_words()        # только словарные токены с метаданными
        -> count_patterns_strict()      # подсчёт вхождений паттернов
        -> write_context_logs()         # сохранение контекстов в CSV
"""

import os
import re
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd
from natasha import Doc

from src.preprocessing import TextPreprocessor

# Символы дефиса и тире в различных Unicode-представлениях
HYPHEN_CHARS = {'-', '‑', '–', '—', '−'}

# Регулярное выражение для проверки наличия буквы в токене
ALPHA_RE = re.compile(r'[A-Za-zА-Яа-яЁё]')

# Части речи, которые считаются служебными (не словарными)
_BANNED_POS = {'PRON', 'NUM', 'ADP', 'SCONJ', 'CCONJ', 'PART', 'INTJ', 'PUNCT', 'SYM'}


# =============================================================================
# Токенизация
# =============================================================================

def tokens_with_punct(text: str, preprocessor: TextPreprocessor) -> list[dict]:
    """
    Токенизирует текст с лемматизацией, сохраняя пунктуацию и позиции.

    Каждый токен — словарь с полями:
        text          : str   — оригинальный текст токена
        pos           : str   — часть речи (Natasha Universal Dependencies)
        lemma         : str | None — нормализованная лемма (None для служебных)
        has_inner_hyphen : bool — содержит ли токен дефис/тире внутри
        start         : int   — начальная позиция в исходном тексте (символы)
        end           : int   — конечная позиция в исходном тексте (символы)

    Args:
        text: исходный текст (сырой, без предобработки)
        preprocessor: экземпляр TextPreprocessor для нормализации

    Returns:
        Список токенов-словарей
    """
    # Нормализуем текст так же, как при предобработке,
    # но НЕ удаляем пунктуацию — она нужна для определения границ фраз
    text = preprocessor.remove_accents(text)
    text = preprocessor.replace_letters(text)
    text = re.sub(r'[\u00AD\u200B\u200C\u200D\u2060]', '', text)

    doc = Doc(text)
    doc.segment(preprocessor.segmenter)
    doc.tag_morph(preprocessor.morph_tagger)

    tokens = []
    for token in doc.tokens:
        pos = token.pos or ''
        token.lemmatize(preprocessor.morph_vocab)

        lemma_raw = (token.lemma or '').replace('ё', 'е')
        corrected = preprocessor.correct_lemma(token.text, lemma_raw, [])
        corrected = corrected.replace('ё', 'е')

        # Токен считается словарным, если содержит букву и не является служебным
        wordlike = bool(ALPHA_RE.search(token.text))
        is_word = wordlike and (pos not in _BANNED_POS)

        lemma_norm = None
        if is_word:
            lemma_lower = corrected.lower()
            lemma_norm = preprocessor.replacements.get(lemma_lower, lemma_lower)

        has_inner_hyphen = any(h in token.text for h in HYPHEN_CHARS)

        tokens.append({
            'text': token.text,
            'pos': pos,
            'lemma': lemma_norm,
            'has_inner_hyphen': has_inner_hyphen,
            'start': token.start,
            'end': token.stop,
        })

    return tokens


def build_indexed_words(tokens: list[dict]) -> list[dict]:
    """
    Извлекает из списка токенов только словарные (с леммой),
    добавляя информацию о соседних дефисах.

    Каждое слово — словарь с полями:
        lemma       : str  — нормализованная лемма
        tok_idx     : int  — индекс в исходном списке tokens
        hyph_left   : bool — слева стоит дефис/тире
        hyph_right  : bool — справа стоит дефис/тире
        inner_hyphen: bool — внутри токена есть дефис/тире
        start       : int  — начальная позиция в тексте
        end         : int  — конечная позиция в тексте

    Args:
        tokens: список токенов из tokens_with_punct()

    Returns:
        Список словарных токенов с метаданными
    """
    words = []
    for i, t in enumerate(tokens):
        if t['lemma'] is None:
            continue

        prev_is_hyphen = (
            i > 0
            and tokens[i - 1]['pos'] == 'PUNCT'
            and tokens[i - 1]['text'] in HYPHEN_CHARS
        )
        next_is_hyphen = (
            i + 1 < len(tokens)
            and tokens[i + 1]['pos'] == 'PUNCT'
            and tokens[i + 1]['text'] in HYPHEN_CHARS
        )

        words.append({
            'lemma': t['lemma'],
            'tok_idx': i,
            'hyph_left': prev_is_hyphen,
            'hyph_right': next_is_hyphen,
            'inner_hyphen': t['has_inner_hyphen'],
            'start': t['start'],
            'end': t['end'],
        })

    return words


# =============================================================================
# Нормализация термина
# =============================================================================

def normalize_term_to_pattern(
    term_str: str,
    preprocessor: TextPreprocessor
) -> tuple[list[str], str]:
    """
    Преобразует термин в паттерн лемм для последующего поиска в тексте.

    Например:
        "биологическая эволюция" -> (["биологический", "эволюция"], "биологический эволюция")
        "яйцеклетка"             -> (["яйцеклетка"], "яйцеклетка")

    Args:
        term_str: термин в исходной форме (из словаря)
        preprocessor: экземпляр TextPreprocessor

    Returns:
        Кортеж (список лемм, паттерн через пробел)
    """
    tks = tokens_with_punct(str(term_str), preprocessor)
    words = build_indexed_words(tks)
    lemmas = [w['lemma'] for w in words]
    pattern = " ".join(lemmas)
    return lemmas, pattern


# =============================================================================
# Подсчёт вхождений паттернов
# =============================================================================

def count_patterns_strict(
    tokens: list[dict],
    patterns_by_len: dict[int, list[str]],
    allow_punct_between: bool = False,
    disallow_hyphen_adjacent_for_unigrams: bool = True,
    collect_patterns: set | None = None,
    context_window: int = 6,
    max_collect_per_pattern: int = 10,
) -> tuple[dict[str, int], dict[str, list[dict]]]:
    """
    Подсчитывает вхождения паттернов (n-грамм лемм) в тексте.

    Алгоритм:
    - Для каждой позиции в списке словарных токенов проверяет,
      образуют ли k последовательных токенов искомый паттерн.
    - Для k=1 (унигаммы) дополнительно проверяет окружение на дефисы.
    - Для k>=2 требует, чтобы слова фразы шли подряд в исходном
      списке токенов (без вклинившихся слов; пунктуация допускается
      при allow_punct_between=True).

    Args:
        tokens: список токенов из tokens_with_punct()
        patterns_by_len: словарь {длина_паттерна: [паттерн, ...]}
            Пример: {1: ["клетка", "ядро"], 2: ["клеточный мембрана"]}
        allow_punct_between: если True — между словами фразы допускается
            только пунктуация (не другие слова)
        disallow_hyphen_adjacent_for_unigrams: если True — унигаммы
            рядом с дефисом/тире не засчитываются
        collect_patterns: множество паттернов, для которых собирать контексты
        context_window: количество токенов слева и справа от совпадения
        max_collect_per_pattern: максимум контекстов на один паттерн

    Returns:
        Кортеж:
            counts  : dict {паттерн: количество_вхождений}
            collected: dict {паттерн: [{"surface", "context",
                                        "start_tok", "end_tok"}, ...]}
    """
    words = build_indexed_words(tokens)
    result: dict[str, int] = {}
    collected: dict[str, list[dict]] = defaultdict(list)
    collect_patterns = set(collect_patterns or [])

    patsets = {k: set(v) for k, v in patterns_by_len.items()}

    # -----------------------------------------------------------------
    # Унигаммы (k = 1)
    # -----------------------------------------------------------------
    if 1 in patsets:
        counts1: Counter = Counter()

        for w in words:
            lemma = w['lemma']
            if lemma not in patsets[1]:
                continue

            # Пропускаем, если рядом дефис/тире (настраивается)
            if disallow_hyphen_adjacent_for_unigrams and (
                w['hyph_left'] or w['hyph_right'] or w['inner_hyphen']
            ):
                continue

            counts1[lemma] += 1

            # Сбор контекста
            if lemma in collect_patterns and len(collected[lemma]) < max_collect_per_pattern:
                collected[lemma].append(
                    _build_context(tokens, w['tok_idx'], w['tok_idx'], context_window)
                )

        for p in patterns_by_len[1]:
            result[p] = counts1.get(p, 0)

    # -----------------------------------------------------------------
    # k-граммы (k >= 2)
    # -----------------------------------------------------------------
    for k, pats in patterns_by_len.items():
        if k == 1:
            continue
        if not pats or len(words) < k:
            for p in pats:
                result[p] = 0
            continue

        counts_k: Counter = Counter()
        patset = patsets[k]

        for i in range(len(words) - k + 1):
            if not _words_are_adjacent(words, tokens, i, k, allow_punct_between):
                continue

            phrase = " ".join(w['lemma'] for w in words[i:i + k])
            if phrase not in patset:
                continue

            counts_k[phrase] += 1

            if phrase in collect_patterns and len(collected[phrase]) < max_collect_per_pattern:
                left_tok = words[i]['tok_idx']
                right_tok = words[i + k - 1]['tok_idx']
                collected[phrase].append(
                    _build_context(tokens, left_tok, right_tok, context_window)
                )

        for p in pats:
            result[p] = counts_k.get(p, 0)

    # Гарантируем нули для всех паттернов, которые не встретились
    for pats in patterns_by_len.values():
        for p in pats:
            result.setdefault(p, 0)

    return result, collected


def _words_are_adjacent(
    words: list[dict],
    tokens: list[dict],
    start: int,
    k: int,
    allow_punct_between: bool,
) -> bool:
    """
    Проверяет, идут ли k слов начиная с позиции start подряд.

    Строгий режим (allow_punct_between=False):
        Индексы токенов должны быть строго последовательными.
    Мягкий режим (allow_punct_between=True):
        Между словами допускается только пунктуация.

    Args:
        words: список словарных токенов
        tokens: полный список токенов (включая пунктуацию)
        start: начальная позиция в words
        k: длина проверяемой фразы
        allow_punct_between: режим проверки

    Returns:
        True, если слова считаются смежными
    """
    for r in range(k - 1):
        idx_a = words[start + r]['tok_idx']
        idx_b = words[start + r + 1]['tok_idx']

        # Строгий режим: токены должны идти подряд
        if idx_b == idx_a + 1:
            continue

        if not allow_punct_between:
            return False

        # Мягкий режим: между токенами только пунктуация
        if idx_b <= idx_a:
            return False
        for t_idx in range(idx_a + 1, idx_b):
            if tokens[t_idx]['pos'] != 'PUNCT':
                return False

    return True


def _build_context(
    tokens: list[dict],
    left_tok: int,
    right_tok: int,
    window: int,
) -> dict:
    """
    Формирует запись о найденном совпадении с контекстом.

    Args:
        tokens: полный список токенов
        left_tok: индекс первого токена совпадения
        right_tok: индекс последнего токена совпадения
        window: размер окна контекста (токенов с каждой стороны)

    Returns:
        Словарь с полями surface, context, start_tok, end_tok
    """
    L = max(0, left_tok - window)
    R = min(len(tokens) - 1, right_tok + window)
    surface = " ".join(t['text'] for t in tokens[left_tok:right_tok + 1])
    context = " ".join(t['text'] for t in tokens[L:R + 1])
    return {
        'surface': surface,
        'context': context,
        'start_tok': left_tok,
        'end_tok': right_tok,
    }


# =============================================================================
# Логирование контекстов
# =============================================================================

def write_context_logs(
    grade: int,
    collected: dict[str, list[dict]],
    pattern_to_terms: dict[str, list[str]],
    out_dir: str | Path,
    docx_path: str | Path,
) -> None:
    """
    Сохраняет контексты найденных терминов в CSV-файл.

    Файл сохраняется в out_dir/grade_{grade}_contexts.csv.
    Если для данного класса не найдено ни одного контекста — файл не создаётся.

    Столбцы CSV:
        Grade       : номер класса
        Doc         : имя файла учебника
        Term        : оригинальный термин (или несколько через |)
        Pattern     : нормализованный паттерн лемм
        Match#      : порядковый номер совпадения
        Surface     : как термин выглядит в тексте
        Context     : контекст вокруг совпадения
        StartTokIdx : индекс первого токена совпадения
        EndTokIdx   : индекс последнего токена совпадения

    Args:
        grade: номер класса (5–11)
        collected: словарь {паттерн: [контексты]}
        pattern_to_terms: словарь {паттерн: [оригинальные термины]}
        out_dir: директория для сохранения
        docx_path: путь к обрабатываемому учебнику (для поля Doc)
    """
    if not collected:
        return

    os.makedirs(out_dir, exist_ok=True)
    rows = []

    for pattern, hits in collected.items():
        terms = " | ".join(pattern_to_terms.get(pattern, [pattern]))
        for j, h in enumerate(hits, start=1):
            rows.append({
                'Grade': grade,
                'Doc': os.path.basename(docx_path),
                'Term': terms,
                'Pattern': pattern,
                'Match#': j,
                'Surface': h['surface'],
                'Context': h['context'],
                'StartTokIdx': h['start_tok'],
                'EndTokIdx': h['end_tok'],
            })

    if not rows:
        return

    csv_path = Path(out_dir) / f"grade_{grade}_contexts.csv"
    pd.DataFrame(rows).to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"[ЛОГ] Контексты сохранены: {csv_path}")


# =============================================================================
# Отладка отдельного термина
# =============================================================================

def debug_term_in_text(
    term_str: str,
    text: str,
    preprocessor: TextPreprocessor,
    allow_punct_between: bool = False,
    disallow_hyphen_adjacent_for_unigrams: bool = False,
    window: int = 6,
) -> None:
    """
    Отладочный инструмент: проверяет, почему термин находится
    или не находится в конкретном тексте.

    Выводит в консоль:
    - нормализованный паттерн термина
    - все токены текста, содержащие подстроку из термина
    - результат подсчёта и найденные контексты

    Args:
        term_str: термин для проверки
        text: текст, в котором ищем
        preprocessor: экземпляр TextPreprocessor
        allow_punct_between: передаётся в count_patterns_strict
        disallow_hyphen_adjacent_for_unigrams: передаётся в count_patterns_strict
        window: размер окна контекста
    """
    separator = "=" * 70

    print(separator)
    print(f"[DEBUG] Термин: {term_str!r}")

    t_lemmas, t_pattern = normalize_term_to_pattern(term_str, preprocessor)
    print(f"[DEBUG] Паттерн: {t_pattern!r} | длина={len(t_lemmas)}")

    # Проверяем, не попал ли термин в таблицу замен
    repl = preprocessor.replacements.get(term_str.lower())
    if repl:
        print(f"[DEBUG] Термин есть в replacements: {term_str.lower()!r} -> {repl!r}")

    # Токенизация текста
    tokens = tokens_with_punct(text, preprocessor)

    # Ищем токены, связанные с первым словом паттерна
    search_substr = t_lemmas[0][:5] if t_lemmas else ''
    print(f"\n[DEBUG] Токены с подстрокой {search_substr!r}:")

    hit_idxs = [
        i for i, t in enumerate(tokens)
        if search_substr in t['text'].lower()
        or search_substr in (t.get('lemma') or '').lower()
    ]

    if not hit_idxs:
        print(f"  Не найдено ни одного токена с подстрокой {search_substr!r}")
    else:
        for i in hit_idxs:
            L = max(0, i - window)
            R = min(len(tokens) - 1, i + window)
            ctx = " ".join(t['text'] for t in tokens[L:R + 1])
            print(
                f"  tok[{i}] text={tokens[i]['text']!r} "
                f"lemma={tokens[i]['lemma']!r} "
                f"pos={tokens[i]['pos']}"
            )
            print(f"    контекст: {ctx}")

    # Подсчёт только по этому термину
    if not t_pattern:
        print("\n[DEBUG] Паттерн пустой — термин не пройдёт фильтры препроцессора.")
        print(separator)
        return

    patterns_local = defaultdict(list)
    patterns_local[len(t_lemmas)].append(t_pattern)

    print(f"\n[DEBUG] Подсчёт вхождений:")
    counts, collected = count_patterns_strict(
        tokens,
        patterns_local,
        allow_punct_between=allow_punct_between,
        disallow_hyphen_adjacent_for_unigrams=disallow_hyphen_adjacent_for_unigrams,
        collect_patterns={t_pattern},
        context_window=window,
        max_collect_per_pattern=10,
    )

    count = counts.get(t_pattern, 0)
    print(f"  Найдено вхождений: {count}")

    if collected.get(t_pattern):
        print("  Контексты:")
        for j, h in enumerate(collected[t_pattern], 1):
            print(f"    #{j}: {h['surface']!r}")
            print(f"         {h['context']!r}")
    else:
        print("  Контексты не собраны — совпадений не было.")

    print(separator)



def merge_spans(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """
    Объединяет перекрывающиеся или смежные интервалы.

    Пример:
        [(0, 5), (3, 8), (10, 15)] -> [(0, 8), (10, 15)]

    Args:
        spans: список интервалов [(start, end), ...]

    Returns:
        Отсортированный список неперекрывающихся интервалов
    """
    if not spans:
        return []

    spans = sorted(spans, key=lambda x: (x[0], x[1]))
    merged = [spans[0]]

    for start, end in spans[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end:
            # Интервалы пересекаются — расширяем последний
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))

    return merged