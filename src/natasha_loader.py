# -*- coding: utf-8 -*-
"""Загрузка natasha без pymorphy2.

В проекте вся морфология — **pymorphy3**. natasha используется только как
контекстный теггер частей речи (``Segmenter`` + ``NewsMorphTagger``), её
собственный лемматизатор ``MorphVocab`` не применяется.

Беда в том, что ``natasha/__init__.py`` безусловно импортирует ``MorphVocab``,
а тот на уровне модуля тянет ``pymorphy2``. Чтобы не держать в окружении
устаревший pymorphy2 (он же ломается на Python 3.11+), подставляем заглушку:
имена ``Parse`` и ``MorphAnalyzer`` нужны только для успешного импорта, а
вызывать их некому — ``MorphVocab()`` мы не создаём.
"""

from __future__ import annotations

import sys
import types


class _StubUnavailable(RuntimeError):
    """Кто-то попытался воспользоваться pymorphy2 — этого быть не должно."""


def _install_pymorphy2_stub() -> None:
    if "pymorphy2" in sys.modules:
        return
    try:
        import pymorphy2  # noqa: F401  — если вдруг реально установлен

        return
    except ImportError:
        pass

    def _fail(*args, **kwargs):
        raise _StubUnavailable(
            "MorphVocab из natasha отключён: в проекте используется pymorphy3"
        )

    class _Unavailable:
        # natasha на этапе определения класса MorphVocab оборачивает
        # PymorphyAnalyzer.parse в lru_cache, поэтому атрибут должен
        # существовать; вызвать его никто не может — MorphVocab не создаётся.
        __init__ = _fail
        parse = _fail
        __call__ = _fail

    class _BaseAnalyzerUnit:
        """Заглушка для yargy.morph.pymorphy2_311_hotfix()."""

    package = types.ModuleType("pymorphy2")
    package.__path__ = []  # чтобы считался пакетом и допускал подмодули

    analyzer = types.ModuleType("pymorphy2.analyzer")
    analyzer.Parse = _Unavailable
    analyzer.MorphAnalyzer = _Unavailable

    units = types.ModuleType("pymorphy2.units")
    units.__path__ = []
    units_base = types.ModuleType("pymorphy2.units.base")
    units_base.BaseAnalyzerUnit = _BaseAnalyzerUnit
    units.base = units_base

    package.analyzer = analyzer
    package.units = units
    package.MorphAnalyzer = _Unavailable

    sys.modules["pymorphy2"] = package
    sys.modules["pymorphy2.analyzer"] = analyzer
    sys.modules["pymorphy2.units"] = units
    sys.modules["pymorphy2.units.base"] = units_base


def load():
    """Возвращает (Doc, Segmenter, NewsEmbedding, NewsMorphTagger)."""
    _install_pymorphy2_stub()

    from natasha.doc import Doc
    from natasha.segment import Segmenter
    from natasha.emb import NewsEmbedding
    from natasha.morph.tagger import NewsMorphTagger

    return Doc, Segmenter, NewsEmbedding, NewsMorphTagger
