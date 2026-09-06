# -*- coding: utf-8 -*-
"""Тесты поиска на реальных случаях из списка единиц культурного кода.

Запуск:  python -m pytest tests -q
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import analyzer, disambiguation  # noqa: E402
from src.lemmatizers import get_backend  # noqa: E402
from src.units import UnitEntry, build_variants  # noqa: E402

BACKENDS = ["pymorphy", "natasha"]


def make_units(names: list[str], backend, group: str = "тест") -> list[UnitEntry]:
    units = [UnitEntry(name=n, group_name=group, group_sheet=group) for n in names]
    build_variants(units, backend)
    return units


def count(names: list[str], text: str, backend_name: str, group: str = "тест") -> dict[str, int]:
    """Сколько раз каждая единица засчитана в тексте."""
    backend = get_backend(backend_name)
    units = make_units(names, backend, group)
    index = analyzer.build_pattern_index(units)
    first = analyzer.first_lemma_set(index)

    tokens = backend.tokenize(text)
    matches = analyzer.find_matches(tokens, index, first, 0)
    disambiguation.resolve(matches, [tokens], units)

    result = {n: 0 for n in names}
    for m in matches:
        if m.status in disambiguation.COUNTED:
            result[units[m.unit_idx].name] += 1
    return result


@pytest.mark.parametrize("backend", BACKENDS)
class TestMatching:
    def test_inflected_multiword(self, backend):
        """Словоизменение внутри многословной единицы."""
        got = count(
            ["Азовское море", "Каспийское море"],
            "К востоку от Каспийского моря и к северу от Азовского моря.",
            backend,
        )
        assert got["Азовское море"] == 1
        assert got["Каспийское море"] == 1

    def test_proverb_with_comma(self, backend):
        """Пунктуация внутри пословицы не мешает совпадению."""
        got = count(
            ["Не имей сто рублей, а имей сто друзей"],
            "Как говорится, не имей сто рублей, а имей сто друзей.",
            backend,
        )
        assert got["Не имей сто рублей, а имей сто друзей"] == 1

    def test_proverb_with_dash(self, backend):
        got = count(
            ["Сделал дело — гуляй смело"],
            "Народная мудрость: сделал дело — гуляй смело.",
            backend,
        )
        assert got["Сделал дело — гуляй смело"] == 1

    def test_function_words_kept(self, backend):
        """Служебные слова входят в паттерн, а не выбрасываются."""
        backend_obj = get_backend(backend)
        units = make_units(["Не имей сто рублей, а имей сто друзей"], backend_obj)
        pattern = units[0].variants[0].pattern
        assert "а" in pattern and "сто" in pattern

    def test_hyphenated(self, backend):
        got = count(
            ["Ванька-встанька", "Царь-бомба"],
            "На полке стоял ванька-встанька, а в музее — царь-бомба.",
            backend,
        )
        assert got["Ванька-встанька"] == 1
        assert got["Царь-бомба"] == 1

    def test_alphanumeric(self, backend):
        got = count(
            ["АК-47", "Восток-1"],
            "Автомат АК-47 и корабль «Восток-1» вошли в историю.",
            backend,
        )
        assert got["АК-47"] == 1
        assert got["Восток-1"] == 1

    def test_latin_homoglyphs(self, backend):
        """В списке «Умa пaлaтa» набрано с латинскими a."""
        got = count(["Умa пaлaтa"], "Про него говорили: ума палата.", backend)
        assert got["Умa пaлaтa"] == 1

    def test_quoted_variant(self, backend):
        """«Крейсер «Аврора»» ловится и по короткому варианту."""
        got = count(
            ["Крейсер «Аврора»"],
            "Выстрел «Авроры» стал сигналом.",
            backend,
        )
        assert got["Крейсер «Аврора»"] == 1

    def test_yo_normalization(self, backend):
        got = count(["Василёк"], "В поле рос василек.", backend)
        assert got["Василёк"] == 1


@pytest.mark.parametrize("backend", BACKENDS)
class TestCaseDisambiguation:
    def test_common_noun_rejected(self, backend):
        """Строчное написание не засчитывается для имени собственного."""
        got = count(["Орёл"], "На гербе изображён орёл с двумя головами.", backend)
        assert got["Орёл"] == 0

    def test_geo_accepted(self, backend):
        got = count(["Орёл"], "Битва развернулась в городе Орёл.", backend)
        assert got["Орёл"] == 1

    def test_common_noun_unit_ignores_case(self, backend):
        """«Храм», «Баня», «Василёк» — нарицательные, регистр им не важен.

        Регрессия: в списке ВСЕ наименования с заглавной буквы, и если
        применять правило регистра ко всем однословным, такие единицы
        перестают находиться вовсе.
        """
        got = count(
            ["Храм", "Баня", "Медведь"],
            "У реки стояла баня, рядом храм, а в лесу бродил медведь.",
            backend,
        )
        assert got == {"Храм": 1, "Баня": 1, "Медведь": 1}

    def test_sentence_initial_needs_support(self, backend):
        """Заглавная в начале предложения сама по себе не доказательство."""
        got = count(["Орёл"], "Орёл кружил высоко над степью.", backend)
        assert got["Орёл"] == 0

    def test_numbered_heading_is_sentence_initial(self, backend):
        """Заголовок «2 Орёл.» — тоже вынужденная заглавная."""
        got = count(["Орёл"], "2 Орёл. Птица эта селится в горах.", backend)
        assert got["Орёл"] == 0

    def test_mid_sentence_capital_is_confident(self, backend):
        got = count(["Орёл"], "Бои шли за город Орёл всё лето.", backend)
        assert got["Орёл"] == 1

    def test_curated_proper_unit(self, backend):
        """«Мы» — роман Замятина; без правила регистра ловил бы каждое «мы»."""
        got = count(["Мы"], "Мы шли долго, и мы очень устали к вечеру.", backend)
        assert got["Мы"] == 0

    def test_curated_proper_unit_with_generic_head(self, backend):
        """Рядом с родовым словом «роман» то же «Мы» уже засчитывается."""
        got = count(["Мы"], "Мы — роман Евгения Замятина.", backend)
        assert got["Мы"] == 1

    def test_multi_lemma_matching(self, backend):
        """«Авроры» pymorphy разбирает первым вариантом как «аврор»."""
        got = count(["Аврора"], "Выстрел «Авроры» стал сигналом.", backend)
        assert got["Аврора"] == 1

    def test_document_rule(self, backend):
        """Одно уверенное вхождение подтверждает спорные в том же тексте."""
        got = count(
            ["Победа"],
            "Автомобиль «Победа» выпускался в Горьком. Победа сошла с конвейера в 1946 году.",
            backend,
        )
        assert got["Победа"] == 2

    def test_generic_head_nearby(self, backend):
        got = count(
            ["Аврора"],
            "Крейсер Аврора стоит на вечной стоянке.",
            backend,
        )
        assert got["Аврора"] == 1

    def test_multiword_ignores_case(self, backend):
        """Правило регистра к многословным единицам не применяется."""
        got = count(["Яблоко раздора"], "Это стало яблоком раздора.", backend)
        assert got["Яблоко раздора"] == 1


LIT = "13. Худ. лит."


@pytest.mark.parametrize("backend", BACKENDS)
class TestTitleCase:
    """Многословные названия произведений требуют заглавной или кавычек."""

    def test_lowercase_phrase_rejected(self, backend):
        got = count(
            ["Война и мир"],
            "Он решает вопросы о войне и мире.",
            backend,
            group=LIT,
        )
        assert got["Война и мир"] == 0

    def test_three_sisters_rejected(self, backend):
        got = count(
            ["Три сестры"],
            "Судьбу человека определяют три сестры — мойры.",
            backend,
            group=LIT,
        )
        assert got["Три сестры"] == 0

    def test_capitalized_title_counted(self, backend):
        got = count(
            ["Война и мир"],
            "В школе изучают Войну и мир целую четверть.",
            backend,
            group=LIT,
        )
        assert got["Война и мир"] == 1

    def test_quoted_title_counted(self, backend):
        got = count(
            ["Война и мир"],
            "Роман «война и мир» переведён на десятки языков.",
            backend,
            group=LIT,
        )
        assert got["Война и мир"] == 1

    def test_proverbs_untouched(self, backend):
        """Пословицы и идиомы правило регистра не затрагивает."""
        got = count(
            ["Яблоко раздора"],
            "Это стало яблоком раздора между ними.",
            backend,
            group="4. Послов.",
        )
        assert got["Яблоко раздора"] == 1


@pytest.mark.parametrize("backend", BACKENDS)
class TestSurnameCollisions:
    """Фамилии совпадают с род. падежом мн. числа обычных слов."""

    def test_surname_not_counted(self, backend):
        got = count(
            ["Соловей", "Пирог", "Ворон", "Волк"],
            "Историк С. Соловьёв, хирург Н. Пирогов, поэт Ю. Воронов и художник Е. Волков.",
            backend,
            group="20. Фауна",
        )
        assert got == {"Соловей": 0, "Пирог": 0, "Ворон": 0, "Волк": 0}

    def test_lowercase_still_counted(self, backend):
        got = count(
            ["Соловей", "Волк"],
            "В роще было много соловьёв, а в лесу волков.",
            backend,
            group="20. Фауна",
        )
        assert got == {"Соловей": 1, "Волк": 1}

    def test_proper_unit_survives(self, backend):
        """У «Авроры» нужная лемма приходит из разбора имени собственного."""
        got = count(["Аврора"], "Выстрел «Авроры» стал сигналом.", backend)
        assert got["Аврора"] == 1


@pytest.mark.parametrize("backend", BACKENDS)
class TestCommonNounGroups:
    """В «природно нарицательных» группах правило регистра не применяется."""

    def test_common_groups_ignore_case(self, backend):
        got = count(
            ["Поле", "Орел"],
            "За околицей было поле, а над ним кружил орёл.",
            backend,
            group="20. Фауна",
        )
        assert got == {"Поле": 1, "Орел": 1}

    def test_flax_found(self, backend):
        got = count(["Лен"], "Крестьяне сеяли лён и коноплю.", backend, group="18. Флора")
        assert got["Лен"] == 1

    def test_city_keeps_case_rule(self, backend):
        """А в «Городах» тот же «Орёл» правило регистра сохраняет."""
        got = count(["Орёл"], "Над степью кружил орёл.", backend, group="22. Города")
        assert got["Орёл"] == 0


@pytest.mark.parametrize("backend", BACKENDS)
class TestGenericHeadVariant:
    def test_fortress_without_head(self, backend):
        got = count(
            ["Крепость Порт-Артур"],
            "Японцы придвинули к Порт-Артуру осадную артиллерию.",
            backend,
        )
        assert got["Крепость Порт-Артур"] == 1

    def test_genitive_modifier_not_stripped(self, backend):
        """«Успенский собор Московского Кремля» не должен ловить «Кремль»."""
        got = count(
            ["Успенский собор Московского Кремля"],
            "Экспонаты хранятся в музеях Московского Кремля.",
            backend,
        )
        assert got["Успенский собор Московского Кремля"] == 0


@pytest.mark.parametrize("backend", BACKENDS)
class TestAbbreviations:
    """Точка после инициала или сокращения — не конец предложения."""

    def test_initial_before_surname(self, backend):
        got = count(
            ["Воробей"],
            "«Незнайка». Худ. фильм, реж. В. Воробьёв, 1976 г. (СССР).",
            backend,
            group="20. Фауна",
        )
        assert got["Воробей"] == 0

    def test_real_sentence_start_still_works(self, backend):
        """После настоящего конца предложения заглавная ничего не значит."""
        got = count(
            ["Орёл"],
            "Стояла тишина. Орёл кружил над степью.",
            backend,
            group="22. Города",
        )
        assert got["Орёл"] == 0


@pytest.mark.parametrize("backend", BACKENDS)
class TestOwnLemmaVariants:
    """Паттерн единицы строится по одному разбору, а он бывает неверным."""

    def test_toponym_not_treated_as_surname(self, backend):
        """«Хохлома»: вероятнее всего разбирается как «хохлом», а не «хохлома»."""
        got = count(
            ["Хохлома"],
            "Промысел удалось сохранить в Хохломе, где издревле изготавливали посуду.",
            backend,
            group="19. Нар. ис-во",
        )
        assert got["Хохлома"] == 1

    def test_surname_still_rejected(self, backend):
        got = count(
            ["Заяц"],
            "Среди писателей-эмигрантов были И. Шмелёв и Б. Зайцев.",
            backend,
            group="20. Фауна",
        )
        assert got["Заяц"] == 0
