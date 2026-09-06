# -*- coding: utf-8 -*-
"""Разрешение спорных вхождений однословных имён собственных.

Задача: «Орёл» (город) не должен ловить птицу орла, «Мороз», «Победа»,
«Союз», «Лада», «Газ» — не должны ловить нарицательные.

Правило применяется ТОЛЬКО к однословным наименованиям, про которые есть
подтверждение, что это имя собственное: помета словаря (Geox/Name/Surn/
Patr/Orgn/Trad) либо ручной список ``data/proper_units.txt``. Одной заглавной
буквы в исходном списке недостаточно — там с заглавной записаны вообще все
наименования, включая «Баня», «Василёк», «Храм», «Царь». Такие нарицательные
единицы, как и все многословные, ищутся без ограничений по регистру.

Классификация вхождения:

* строчная буква в тексте          -> ``rejected_lowercase`` (в счёт не идёт)
* заглавная в свободной позиции    -> ``confident``
* заглавная в вынужденной позиции  -> спорное, решается каскадом:

  1. у словоформы вообще нет нарицательного прочтения (только Geox/Name/…) ->
     принять. Для омонимов вроде «Орёл» это правило молчит: помета Geox
     стоит и у города, и у птицы, поэтому ничего не доказывает;
  2. слово в кавычках («Победа», «Союз») -> принять;
  3. рядом (±3 словарных токена) родовое слово из data/generic_heads.txt
     (крейсер, ледокол, город, ракета, …) -> принять;
  4. у этой же единицы в этом же учебнике есть хотя бы одно уверенное
     вхождение -> принять;
  5. иначе -> отклонить.

Отклонённые вхождения не пропадают: они попадают на лист «Спорные» с
контекстом и причиной, а в docx подсвечиваются оранжевым.
"""

from __future__ import annotations

from collections import defaultdict

import config
from src.lemmatizers import (
    Token,
    word_has_common_reading,
    word_is_proper,
    word_lemmas,
    word_proper_lemmas,
)
from src.units import UnitEntry, is_all_caps, is_capitalized_single_word

STATUS_CONFIDENT = "confident"
STATUS_ACCEPTED = "disputed_accepted"
STATUS_REJECTED = "disputed_rejected"
STATUS_LOWERCASE = "rejected_lowercase"

#: статусы, которые идут в основной счётчик
COUNTED = frozenset({STATUS_CONFIDENT, STATUS_ACCEPTED})

OPEN_QUOTES = "«\"„“'"
CLOSE_QUOTES = "»\"“”'"

_GENERIC_HEADS: set[str] | None = None


def load_generic_heads(path=None) -> set[str]:
    """Читает словарь родовых слов (кэшируется)."""
    global _GENERIC_HEADS
    if _GENERIC_HEADS is None:
        path = path or config.GENERIC_HEADS_FILE
        heads: set[str] = set()
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if line and not line.startswith("#"):
                    heads.add(line.replace("ё", "е").lower())
        _GENERIC_HEADS = heads
    return _GENERIC_HEADS


def _first_alpha_is_upper(text: str) -> bool:
    for ch in text:
        if ch.isalpha():
            return ch.isupper()
    return False


def _is_all_caps_token(text: str) -> bool:
    letters = [ch for ch in text if ch.isalpha()]
    return len(letters) > 1 and all(ch.isupper() for ch in letters)


def _in_quotes(tokens: list[Token], tok_start: int, tok_end: int) -> bool:
    """Обрамлено ли совпадение кавычками."""
    left = tokens[tok_start - 1].text if tok_start > 0 else ""
    right = tokens[tok_end + 1].text if tok_end + 1 < len(tokens) else ""
    return bool(left) and left[-1] in OPEN_QUOTES and bool(right) and right[0] in CLOSE_QUOTES


def _generic_head_nearby(tokens: list[Token], tok_start: int, tok_end: int) -> str | None:
    """Ищет родовое слово в окне ±N словарных токенов; возвращает его лемму."""
    heads = load_generic_heads()
    window = config.GENERIC_HEAD_WINDOW

    seen = 0
    i = tok_start - 1
    while i >= 0 and seen < window:
        if tokens[i].is_word:
            seen += 1
            if tokens[i].lemma in heads:
                return tokens[i].lemma
        i -= 1

    seen = 0
    i = tok_end + 1
    while i < len(tokens) and seen < window:
        if tokens[i].is_word:
            seen += 1
            if tokens[i].lemma in heads:
                return tokens[i].lemma
        i += 1
    return None


_PROPER_UNITS: set[str] | None = None


def load_proper_units(path=None) -> set[str]:
    """Ручной список однословных наименований — имён собственных.

    Нужен там, где морфология не помогает: у «Победа», «Союз», «Москвич»,
    «Спутник», «ГАЗ» в словаре нет пометы имени собственного, хотя в этом
    списке это марки машин, ракет и спутников.
    """
    global _PROPER_UNITS
    if _PROPER_UNITS is None:
        path = path or config.PROPER_UNITS_FILE
        names: set[str] = set()
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if line and not line.startswith("#"):
                    names.add(line.replace("ё", "е").lower())
        _PROPER_UNITS = names
    return _PROPER_UNITS


MODE_PROPER = "proper"
MODE_TITLE = "title"


def case_check_mode(unit: UnitEntry, variant_idx: int) -> str | None:
    """Какое правило регистра применять к варианту наименования.

    * ``MODE_PROPER`` — однословное имя собственное («Орёл», «Мороз», «Мы»);
    * ``MODE_TITLE``  — многословное название произведения из групп
      «Художественная литература», «Кино, театр», «ИЗО»;
    * ``None``        — регистр не важен (нарицательные, пословицы, фольклор).
    """
    if needs_case_check(unit, variant_idx):
        return MODE_PROPER
    variant = unit.variants[variant_idx]
    if len(variant.pattern) > 1 and unit.group_sheet in config.TITLE_GROUPS:
        return MODE_TITLE
    return None


def needs_case_check(unit: UnitEntry, variant_idx: int) -> bool:
    """Применять ли правило регистра к этому варианту наименования.

    В исходном списке ВСЕ наименования записаны с заглавной буквы, поэтому
    одного регистра мало: иначе под правило попали бы «Баня», «Василёк»,
    «Храм», «Царь» — обычные нарицательные, которые в тексте пишутся со
    строчной и должны находиться.

    Требуем дополнительное подтверждение, что это имя собственное:
    помета словаря (Geox/Name/Surn/Patr/Orgn/Trad) либо ручной список
    ``data/proper_units.txt``.
    """
    variant = unit.variants[variant_idx]
    if len(variant.pattern) != 1:
        return False
    if not is_capitalized_single_word(variant.surface):
        return False

    word = variant.surface.strip("«»\"'()")
    # Ручной список сильнее всего остального.
    if word.replace("ё", "е").lower() in load_proper_units():
        return True
    # В «природно нарицательных» группах помете словаря не верим: у «Поле»,
    # «Лен», «Орел» (птица) есть разбор Geox, и правило регистра лишало их
    # всех вхождений со строчной буквы.
    if unit.group_sheet in config.COMMON_NOUN_GROUPS:
        return False
    return word_is_proper(word)


def resolve(matches, paragraphs: list[list[Token]], units: list[UnitEntry]) -> None:
    """Проставляет ``status`` и ``reason`` каждому вхождению (на месте)."""
    load_generic_heads()

    deferred: list = []
    confident_units: set[int] = set()

    for m in matches:
        unit = units[m.unit_idx]
        variant = unit.variants[m.variant_idx]

        if _is_surname_collision(m, paragraphs, variant):
            m.status = STATUS_REJECTED
            m.reason = "слово с заглавной разбирается как имя собственное"
            continue

        mode = case_check_mode(unit, m.variant_idx)
        if mode is None:
            m.status = STATUS_CONFIDENT
            m.reason = ""
            confident_units.add(m.unit_idx)
            continue

        tokens = paragraphs[m.para_idx]
        token = tokens[m.tok_start]

        if mode == MODE_TITLE:
            _resolve_title(m, tokens, confident_units, deferred)
            continue

        if not _first_alpha_is_upper(token.text):
            m.status = STATUS_LOWERCASE
            m.reason = "строчная буква"
            continue

        forced = m.sentence_initial or (
            _is_all_caps_token(token.text) and not is_all_caps(unit.variants[m.variant_idx].surface)
        )
        if not forced:
            m.status = STATUS_CONFIDENT
            m.reason = "заглавная в середине предложения"
            confident_units.add(m.unit_idx)
            continue

        # Вынужденная заглавная — каскад правил.
        # 1. Морфология помогает, только если у словоформы НЕТ нарицательного
        #    прочтения: у «Орёл» помета Geox есть всегда, независимо от того,
        #    город это или птица, — такой сигнал бесполезен. А вот «Москва»
        #    или «Пушкина» разбираются только как имена собственные.
        proper = token.grammemes & config.PROPER_GRAMMEMES
        if proper and not _has_common_reading(token):
            m.status = STATUS_ACCEPTED
            m.reason = "морфология: только " + ", ".join(sorted(proper))
            continue

        if _in_quotes(tokens, m.tok_start, m.tok_end):
            m.status = STATUS_ACCEPTED
            m.reason = "в кавычках"
            continue

        head = _generic_head_nearby(tokens, m.tok_start, m.tok_end)
        if head:
            m.status = STATUS_ACCEPTED
            m.reason = f"родовое слово рядом: {head}"
            continue

        deferred.append(m)

    # Правило документа: единица уже уверенно встречалась в этом учебнике.
    for m in deferred:
        if m.unit_idx in confident_units:
            m.status = STATUS_ACCEPTED
            m.reason = "правило документа: есть уверенное вхождение"
        else:
            m.status = STATUS_REJECTED
            m.reason = "заглавная вынужденная, подтверждений нет"


def _is_surname_collision(m, paragraphs, variant) -> bool:
    """Совпало ли слово единицы с фамилией или топонимом.

    Русские фамилии совпадают с родительным падежом множественного числа
    обычных слов: «Соловьёв» — это и фамилия, и форма слова «соловей»,
    «Пирогов» — «пирог», «Воронов» — «ворон», «Лены» — и «лён», и река Лена.
    Поиск идёт по всем возможным леммам словоформы, поэтому историк
    С. Соловьёв засчитывался как птица.

    Признак ошибки: слово написано с заглавной буквы В СЕРЕДИНЕ предложения,
    у него есть разбор как имени собственного, но лемма единицы через этот
    разбор не получается. У «Авроры» лемма «аврора» приходит как раз из
    разбора Name, поэтому крейсер не страдает; у «Победа» и «АК-47» разбора
    как имени собственного нет вовсе, и правило молчит.

    Строчные написания правило не трогает: «соловьёв» действительно может
    быть птицами. В начале предложения тоже не трогает — там заглавная
    буква обязательна и ничего не значит.
    """
    if len(variant.pattern) != 1:
        return False
    if m.sentence_initial:
        return False

    token = paragraphs[m.para_idx][m.tok_start]
    if not _first_alpha_is_upper(token.text):
        return False

    proper = word_proper_lemmas(token.text)
    if not proper:
        return False

    # Сравниваем со ВСЕМИ леммами самого наименования, а не с одной основной:
    # у «Хохлома» самый вероятный разбор даёт лемму «хохлом», и вхождение
    # «в Хохломе» отбрасывалось как чужая фамилия.
    own = word_lemmas(variant.surface.strip("«»\"'()")) | {variant.pattern[0]}
    return not (own & proper)


def _resolve_title(m, tokens, confident_units: set[int], deferred: list) -> None:
    """Правило регистра для многословных названий произведений.

    Название засчитывается, если оно взято в кавычки либо начинается с
    заглавной буквы. «Война и мир» в обороте «вопросы о войне и мире» и
    «Три сестры» в «судьбу определяют три сестры» отсеиваются.
    """
    if _in_quotes(tokens, m.tok_start, m.tok_end):
        m.status = STATUS_CONFIDENT
        m.reason = "название в кавычках"
        confident_units.add(m.unit_idx)
        return

    if not _first_alpha_is_upper(tokens[m.tok_start].text):
        m.status = STATUS_REJECTED
        m.reason = "название со строчной буквы и без кавычек"
        return

    if not m.sentence_initial:
        m.status = STATUS_CONFIDENT
        m.reason = "название с заглавной в середине предложения"
        confident_units.add(m.unit_idx)
        return

    head = _generic_head_nearby(tokens, m.tok_start, m.tok_end)
    if head:
        m.status = STATUS_ACCEPTED
        m.reason = f"родовое слово рядом: {head}"
        return

    deferred.append(m)


def summarize_statuses(matches) -> dict[int, dict[str, int]]:
    """Счётчики статусов по единицам: {unit_idx: {status: count}}."""
    out: dict[int, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for m in matches:
        out[m.unit_idx][m.status] += 1
    return {k: dict(v) for k, v in out.items()}


def _has_common_reading(token: Token) -> bool:
    """Есть ли у словоформы прочтение как нарицательного слова."""
    return word_has_common_reading(token.text)
