# -*- coding: utf-8 -*-
"""Чтение списка объектов культурного кода и построение лемма-паттернов.

Структура исходной книги ``Список наименований объектов КК.xlsx``:

* лист ``Общий`` — полный список; колонка B содержит название тематической
  группы только в первой строке блока, дальше её надо «протягивать» вниз.
  После строки-маркера «Дополнительно (без группы)» идёт хвост из ~150
  позиций без группы;
* 23 тематических листа (``1. Трад. быт`` … ``23. Кален. празд.``) — те же
  единицы, разбитые по группам.

Результат: список ``UnitEntry`` — по одной записи на пару
«наименование × тематическая группа» (согласовано с заказчиком).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import openpyxl

import config
from src.text_utils import (
    normalize_lemma,
    remove_accents,
    replace_latin_homoglyphs,
    is_wordlike,
)

#: строка-маркер, после которой в листе «Общий» идут единицы без группы
NO_GROUP_MARKER = "Дополнительно (без группы)"

#: название и ключ листа для хвостового блока
NO_GROUP_SHEET = "24. Доп. без группы"
NO_GROUP_NAME = "Дополнительно (без группы)"

QUOTED_RE = re.compile(r"[«\"]([^»\"]+)[»\"]")
PAREN_RE = re.compile(r"\s*\([^)]*\)")
HEAD_QUOTED_RE = re.compile(r"^([^«\"]+)[«\"]([^»\"]+)[»\"]")


@dataclass
class Variant:
    """Один поисковый вариант наименования."""

    surface: str
    pattern: tuple[str, ...]
    kind: str  # full | quoted | noparen | head+quoted

    @property
    def length(self) -> int:
        return len(self.pattern)


@dataclass
class UnitEntry:
    """Пара «наименование × тематическая группа»."""

    name: str
    group_name: str
    group_sheet: str
    variants: list[Variant] = field(default_factory=list)

    @property
    def key(self) -> tuple[str, str]:
        return (self.name, self.group_sheet)


# ---------------------------------------------------------------------------
# Чтение книги
# ---------------------------------------------------------------------------


def _cell(value) -> str:
    """Значение ячейки как строка; ячейки из одних пробелов считаются пустыми."""
    if value is None:
        return ""
    return str(value).strip()


def _canon(name: str) -> str:
    """Ключ для сопоставления одинаковых наименований между листами."""
    text = replace_latin_homoglyphs(remove_accents(name))
    text = re.sub(r"\s+", " ", text).strip()
    return normalize_lemma(text)


def _read_sheet(ws) -> list[tuple[str, str]]:
    """Читает лист как список пар (группа, наименование).

    Группа берётся из колонки B с протяжкой вниз. Для строк после маркера
    «Дополнительно (без группы)» возвращается пустая группа.
    """
    out: list[tuple[str, str]] = []
    group = ""
    after_marker = False
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 3:
            continue
        cell_group = _cell(row[1])
        name = _cell(row[2])
        if cell_group:
            group = cell_group
        if not name:
            continue
        if name == NO_GROUP_MARKER:
            after_marker = True
            continue
        out.append(("" if after_marker else group, name))
    return out


def load_units(path=None) -> list[UnitEntry]:
    """Собирает итоговый список пар «наименование × группа»."""
    path = path or config.UNITS_XLSX
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    general_sheet = wb.sheetnames[0]
    thematic_sheets = wb.sheetnames[1:]

    # 1) тематические листы: группа -> имя листа, и множество наименований
    group_to_sheet: dict[str, str] = {}
    pairs: dict[tuple[str, str], UnitEntry] = {}
    seen_in_thematic: set[str] = set()

    for sheet in thematic_sheets:
        rows = _read_sheet(wb[sheet])
        for group, name in rows:
            if group:
                group_to_sheet.setdefault(group, sheet)
            seen_in_thematic.add(_canon(name))
            key = (_canon(name), sheet)
            if key not in pairs:
                pairs[key] = UnitEntry(
                    name=name, group_name=group or sheet, group_sheet=sheet
                )

    # 2) лист «Общий»: сгруппированная часть + хвост без группы
    for group, name in _read_sheet(wb[general_sheet]):
        canon = _canon(name)
        if group:
            sheet = group_to_sheet.get(group)
            if sheet is None:
                # Группа есть только в «Общем» — заводим отдельный лист.
                sheet = group[:31]
                group_to_sheet[group] = sheet
            key = (canon, sheet)
            if key not in pairs:
                pairs[key] = UnitEntry(name=name, group_name=group, group_sheet=sheet)
        elif canon not in seen_in_thematic:
            key = (canon, NO_GROUP_SHEET)
            if key not in pairs:
                pairs[key] = UnitEntry(
                    name=name, group_name=NO_GROUP_NAME, group_sheet=NO_GROUP_SHEET
                )

    wb.close()

    units = list(pairs.values())
    units.sort(key=lambda u: (_sheet_order(u.group_sheet), _canon(u.name)))
    return units


def _sheet_order(sheet: str) -> tuple[int, str]:
    """Порядок листов: «1. …», «2. …», … «24. Доп. без группы»."""
    m = re.match(r"^(\d+)\.", sheet)
    return (int(m.group(1)) if m else 999, sheet)


def group_sheets(units: list[UnitEntry]) -> list[str]:
    """Имена листов по группам в правильном порядке."""
    seen: dict[str, None] = {}
    for u in sorted(units, key=lambda u: _sheet_order(u.group_sheet)):
        seen.setdefault(u.group_sheet, None)
    return list(seen)


# ---------------------------------------------------------------------------
# Варианты и лемма-паттерны
# ---------------------------------------------------------------------------


#: предлоги: если родовая часть кончается предлогом, это не «родовое слово +
#: имя», а цельное название («Золотые ворота в Киеве»)
PREPOSITIONS = frozenset({
    "в", "во", "на", "у", "к", "ко", "с", "со", "из", "от", "до", "по",
    "за", "над", "под", "при", "про", "для", "об", "о",
})


def strip_generic_head(name: str) -> str | None:
    """Отбрасывает родовое слово перед именем собственным.

    «Крепость Порт-Артур» -> «Порт-Артур», «Замок Ринген» -> «Ринген»,
    «Атомная подводная лодка К-278 Комсомолец» -> «К-278 Комсомолец».

    В учебнике такие объекты почти никогда не названы полностью: пишут
    «атаковали Порт-Артур», а не «атаковали крепость Порт-Артур».

    Не срабатывает там, где заглавная буква — часть цельного названия:
    «Золотые ворота в Киеве» (родовая часть кончается предлогом),
    «Северная Венеция», «Александр Невский» (одно слово перед именем, и оно
    не родовое).
    """
    from src.disambiguation import load_generic_heads

    words = name.split()
    if len(words) < 2:
        return None

    for i in range(1, len(words)):
        word = words[i]
        if not (word[:1].isupper() or any(ch.isdigit() for ch in word)):
            continue

        prefix = words[:i]
        if prefix[-1].lower().strip(".,") in PREPOSITIONS:
            return None
        # внутри родовой части заглавных быть не должно
        if any(w[:1].isupper() for w in prefix[1:]):
            return None

        # последнее слово родовой части должно быть родовым словом
        if _canon(prefix[-1]) not in load_generic_heads():
            return None

        # родительный падеж после родового слова — это определение, а не имя:
        # «Успенский собор Московского Кремля» -> не «Московского Кремля»
        if any(_is_genitive_only(w) for w in words[i:]):
            return None

        rest = " ".join(words[i:])
        return rest if rest != name else None
    return None


def _is_genitive_only(word: str) -> bool:
    """Все разборы словоформы — родительный падеж."""
    if any(ch.isdigit() or ch == "-" for ch in word):
        return False
    from src.lemmatizers import get_morph

    parses = get_morph().parse(word.strip("«»\"'(),."))
    if not parses:
        return False
    return all("gent" in p.tag.grammemes for p in parses)


def make_variant_surfaces(name: str) -> list[tuple[str, str]]:
    """Порождает поисковые варианты наименования: [(строка, тип)].

    * ``full``        — наименование как есть;
    * ``noparen``     — без скобочного пояснения;
    * ``quoted``      — содержимое кавычек («Аврора», «Божье Предвидение»);
    * ``head+quoted`` — родовое слово перед кавычками + содержимое кавычек.

    Кавычки при токенизации всё равно становятся пунктуацией, поэтому
    ``head+quoted`` часто совпадает с ``full``/``noparen`` — дубликаты
    отсекаются позже, при построении паттернов.
    """
    out: list[tuple[str, str]] = [(name, "full")]

    noparen = PAREN_RE.sub("", name).strip()
    if noparen and noparen != name:
        out.append((noparen, "noparen"))

    stripped = strip_generic_head(name)
    if stripped:
        out.append((stripped, "без родового слова"))

    for source in {name, noparen} - {""}:
        for quoted in QUOTED_RE.findall(source):
            quoted = quoted.strip()
            if quoted and quoted != name:
                out.append((quoted, "quoted"))
        m = HEAD_QUOTED_RE.match(source)
        if m:
            head = m.group(1).strip(" -—:,")
            combined = f"{head} {m.group(2).strip()}".strip()
            if head and combined != name:
                out.append((combined, "head+quoted"))
    return out


def surface_to_pattern(surface: str, backend) -> tuple[str, ...]:
    """Превращает строку в кортеж лемм (пунктуация отбрасывается)."""
    tokens = backend.tokenize(remove_accents(surface))
    return tuple(t.lemma for t in tokens if t.is_word)


def build_variants(units: list[UnitEntry], backend) -> None:
    """Заполняет ``UnitEntry.variants`` лемма-паттернами.

    Варианты одной единицы дедуплицируются по паттерну; производные варианты
    (не ``full``) короче ``MIN_VARIANT_LEN`` символов отбрасываются.
    """
    cache: dict[str, tuple[str, ...]] = {}

    for unit in units:
        seen: set[tuple[str, ...]] = set()
        variants: list[Variant] = []
        for surface, kind in make_variant_surfaces(unit.name):
            if kind != "full" and len(surface) < config.MIN_VARIANT_LEN:
                continue
            if surface not in cache:
                cache[surface] = surface_to_pattern(surface, backend)
            pattern = cache[surface]
            if not pattern or pattern in seen:
                continue
            seen.add(pattern)
            variants.append(Variant(surface=surface, pattern=pattern, kind=kind))
        unit.variants = variants


# ---------------------------------------------------------------------------
# Признаки для разрешения спорных вхождений
# ---------------------------------------------------------------------------


def is_capitalized_single_word(surface: str) -> bool:
    """Однословное наименование, записанное с заглавной буквы.

    Именно к таким применяется правило «требовать заглавную в тексте».
    """
    words = [w for w in surface.split() if is_wordlike(w)]
    if len(words) != 1:
        return False
    first = words[0].lstrip("«\"'(")
    return bool(first) and first[0].isupper()


def is_all_caps(surface: str) -> bool:
    """Наименование записано целиком заглавными (ГАЗ, АК-47, ППШ)."""
    letters = [ch for ch in surface if ch.isalpha()]
    return bool(letters) and all(ch.isupper() for ch in letters)
