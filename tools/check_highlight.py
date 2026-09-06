# -*- coding: utf-8 -*-
"""Проверка, что подсветка не испортила документы.

Для каждой пары «исходный учебник — подсвеченная копия» сверяет:

* текст абзацев совпадает символ в символ;
* форматирование каждого символа (``w:rPr`` его run'а без ``w:color``)
  не изменилось — то есть подсветка не съела жирный, курсив, кегль и стили;
* сколько символов покрашено красным и оранжевым.

Запуск:  python tools/check_highlight.py
"""

from __future__ import annotations

import copy
import re
import sys
from pathlib import Path

from docx.oxml.ns import qn
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from src import docio

NS_ATTR = re.compile(r'\sxmlns:[a-z0-9]+="[^"]*"')


def char_profile(view) -> tuple[list[str], list[str]]:
    """Для каждого символа абзаца: (форматирование без цвета, цвет)."""
    formats: list[str] = []
    colors: list[str] = []
    for run in view.runs:
        rpr = run._element.find(qn("w:rPr"))
        if rpr is None:
            signature, color = "", ""
        else:
            clone = copy.deepcopy(rpr)
            color_el = clone.find(qn("w:color"))
            color = color_el.get(qn("w:val")) if color_el is not None else ""
            for el in clone.findall(qn("w:color")):
                clone.remove(el)
            signature = NS_ATTR.sub("", etree.tostring(clone, encoding="unicode"))
        formats.extend([signature] * len(run.text))
        colors.extend([color] * len(run.text))
    return formats, colors


def main() -> int:
    problems = 0

    for section in sorted(p for p in config.OUT_HIGHLIGHT.iterdir() if p.is_dir()):
        print(f"\n### {section.name}")
        for out_path in sorted(section.glob("*.docx")):
            src_path = config.CORPUS_ROOT / section.name / out_path.name
            if not src_path.exists():
                print(f"  {out_path.name}: исходник не найден")
                problems += 1
                continue

            before = docio.paragraph_views(docio.open_document(src_path))
            after = docio.paragraph_views(docio.open_document(out_path))

            if len(before) != len(after):
                print(f"  {out_path.name}: разное число абзацев "
                      f"({len(before)} / {len(after)})")
                problems += 1
                continue

            text_diff = sum(1 for a, b in zip(before, after) if a.text != b.text)
            format_diff = 0
            red = orange = 0
            for a, b in zip(before, after):
                fa, _ = char_profile(a)
                fb, cb = char_profile(b)
                if fa != fb:
                    format_diff += 1
                red += sum(1 for c in cb if c == config.COLOR_CONFIDENT)
                orange += sum(1 for c in cb if c == config.COLOR_DISPUTED)

            status = "OK" if not text_diff and not format_diff else "ПРОБЛЕМА"
            if text_diff or format_diff:
                problems += 1
            print(
                f"  {out_path.name[:48]:48} {status}: абзацев {len(before)}, "
                f"текст расходится в {text_diff}, форматирование в {format_diff}; "
                f"красным {red} симв., оранжевым {orange}"
            )

    print()
    print("Проблем не найдено." if not problems else f"Проблемных файлов: {problems}")
    return 0 if not problems else 1


if __name__ == "__main__":
    raise SystemExit(main())
