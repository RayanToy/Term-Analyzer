# -*- coding: utf-8 -*-
"""Все модули проекта импортируются.

Тесты поиска не трогают ``src/report.py``, ``main.py`` и инструменты, поэтому
синтаксическая ошибка в них раньше обнаруживалась только на полном прогоне —
после двадцати минут вычислений. Этот тест закрывает такую дыру.
"""

from __future__ import annotations

import importlib
import pkgutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

MODULES = ["config", "main"] + [
    f"src.{m.name}" for m in pkgutil.iter_modules([str(ROOT / "src")])
]

TOOLS = sorted(p.stem for p in (ROOT / "tools").glob("*.py") if p.stem != "__init__")


@pytest.mark.parametrize("name", MODULES)
def test_module_imports(name):
    importlib.import_module(name)


@pytest.mark.parametrize("name", TOOLS)
def test_tool_compiles(name):
    """Инструменты только компилируем: импорт запускает тяжёлые зависимости."""
    source = (ROOT / "tools" / f"{name}.py").read_text(encoding="utf-8")
    compile(source, f"tools/{name}.py", "exec")
