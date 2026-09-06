# -*- coding: utf-8 -*-
"""Показывает фрагменты .docx, которые не нашлись в PDF.

`check_pdf_vs_docx.py` сообщает только счётчик таких фрагментов. Здесь они
печатаются целиком вместе с диагностикой: какие редкие слова фрагмента
отсутствуют в словаре PDF и находятся ли они там в изменённом виде
(разорванные переносом, склеенные, с потерянными буквами).

Запуск:  python tools/show_pdf_gaps.py [подстрока имени учебника ...]
"""

from __future__ import annotations

import difflib
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from tools.check_pdf_vs_docx import (
    FRAGMENT_THRESHOLD,
    LONG_WORD,
    NOISE,
    SHINGLE,
    SPOT_CHECK,
    docx_words,
    grade_of,
    normalize,
    pdf_words,
)


def similar_in_pdf(word: str, pdf_vocab: set[str]) -> list[str]:
    """Ищет слово в словаре PDF в изменённом виде."""
    hits = [w for w in pdf_vocab if w.startswith(word[:6]) or word.startswith(w[:6])]
    hits.sort(key=lambda w: -difflib.SequenceMatcher(None, word, w).ratio())
    return hits[:3]


def main(argv: list[str]) -> int:
    wanted = argv or None

    for section in sorted(p for p in config.CORPUS_ROOT.iterdir() if p.is_dir()):
        docx_by_grade = {}
        for path in sorted(section.glob("*.docx")):
            if path.name in config.SKIP_FILES:
                continue
            grade = grade_of(path)
            if grade:
                docx_by_grade[grade] = path

        for pdf in sorted(section.glob("*.pdf")):
            grade = grade_of(pdf)
            docx = docx_by_grade.get(grade or "")
            if docx is None:
                continue
            if wanted and not any(w in pdf.name or w in docx.name for w in wanted):
                continue

            a = normalize(docx_words(docx))
            raw_pdf = [w for w in pdf_words(pdf) if not NOISE.match(w)]
            b = normalize(raw_pdf)
            pdf_vocab = set(b)

            opcodes = difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes()

            fragments = []
            for tag, i1, i2, _j1, _j2 in opcodes:
                if tag != "equal" and i2 - i1 >= SHINGLE * 2:
                    fragments.append((i1, i2))
            fragments.sort(key=lambda p: p[0] - p[1])

            checked = 0
            print(f"\n{'=' * 78}\n### {section.name} / {grade} класс")
            for i1, i2 in fragments:
                words = a[i1:i2]
                rare = [w for w in words if len(w) >= LONG_WORD]
                if len(rare) < 3:
                    continue
                checked += 1
                if checked > SPOT_CHECK:
                    break
                present = [w for w in rare if w in pdf_vocab]
                if len(present) / len(rare) >= FRAGMENT_THRESHOLD:
                    continue

                missing = [w for w in rare if w not in pdf_vocab]
                print(f"\n  Фрагмент ({len(words)} слов, редких {len(rare)}, "
                      f"нет в PDF {len(missing)}):")
                print("    " + " ".join(words)[:400])
                print("    Отсутствующие слова и что похожее есть в PDF:")
                for word in missing[:8]:
                    near = similar_in_pdf(word, pdf_vocab)
                    print(f"      {word:22} -> {near if near else 'ничего похожего'}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
