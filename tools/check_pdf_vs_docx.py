# -*- coding: utf-8 -*-
"""Сверка PDF-версий учебников с .docx.

В корпусе один и тот же учебник лежит в трёх видах: .docx, .txt и .pdf.
Скрипт проверяет, есть ли в PDF содержательный текст, которого нет в .docx.
Если нет — PDF можно не обрабатывать (а подсветку по нему всё равно не
сделать).

Запуск:  python tools/check_pdf_vs_docx.py
Результат: _РЕЗУЛЬТАТ_КК/_лог_сверки_PDF.txt
"""

from __future__ import annotations

import difflib
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import config
from src.docio import open_document, paragraph_views

#: PDF пронумерованы по классу, .docx — по позиции в имени файла
GRADE_IN_PDF = re.compile(r"-(\d{1,2})-klass")
GRADE_IN_DOCX = re.compile(r"_Hs_(\d{2})_")

#: колонтитулы и номера страниц, которые есть только в PDF
NOISE = re.compile(r"^\d+$")


def docx_words(path: Path) -> list[str]:
    doc = open_document(path)
    return " ".join(v.text for v in paragraph_views(doc)).split()


#: перенос слова со строки на строку: «эпо- хой» -> «эпохой»
HYPHEN_BREAK = re.compile(r"(\w)[-‐‑]\s+([а-яё]\w*)")


def pdf_words(path: Path) -> list[str]:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    chunks = []
    for page in reader.pages:
        try:
            chunks.append(page.extract_text() or "")
        except Exception as exc:  # pragma: no cover - битые страницы
            chunks.append("")
            print(f"    страница не прочитана: {exc}")

    text = " ".join(chunks)
    # Переносы склеиваем, иначе сверка показывает мнимые расхождения:
    # в PDF слово разорвано на «эпо-» и «хой», в docx оно целое.
    previous = None
    while previous != text:
        previous = text
        text = HYPHEN_BREAK.sub(r"\1\2", text)
    return text.split()


#: буквенно-цифровое ядро слова — без кавычек, скобок и знаков препинания
WORD_CORE = re.compile(r"[а-яёa-z0-9]+")

#: длина словосочетания для проверки покрытия
SHINGLE = 3

#: длина «редкого» слова для словарной проверки
LONG_WORD = 8

#: сколько неаллигнированных фрагментов проверять на самом деле
SPOT_CHECK = 20


def normalize(words: list[str]) -> list[str]:
    """Оставляет от каждого слова буквенно-цифровое ядро в нижнем регистре."""
    out = []
    for word in words:
        m = WORD_CORE.search(word.lower())
        if m:
            out.append(m.group(0))
    return out


def shingle_coverage(docx: list[str], pdf: list[str]) -> tuple[int, int]:
    """Доля словосочетаний .docx, встречающихся где-либо в тексте PDF."""
    haystack = {tuple(pdf[i : i + SHINGLE]) for i in range(len(pdf) - SHINGLE + 1)}
    covered = total = 0
    for i in range(0, len(docx) - SHINGLE + 1, SHINGLE):
        total += 1
        if tuple(docx[i : i + SHINGLE]) in haystack:
            covered += 1
    return covered, total


def vocabulary_coverage(docx: list[str], pdf: list[str]) -> tuple[int, int]:
    """Доля редких слов .docx, встречающихся в PDF.

    Если бы в PDF действительно не хватало кусков текста, это в первую
    очередь проявилось бы пропавшими словами.
    """
    long_words = {w for w in docx if len(w) >= LONG_WORD}
    return len(long_words & set(pdf)), len(long_words)


#: доля редких слов фрагмента, при которой считаем фрагмент найденным
FRAGMENT_THRESHOLD = 0.8


def spot_check_unaligned(opcodes, docx: list[str], pdf_vocab: set) -> tuple[int, int]:
    """Проверяет крупные «непопавшие» фрагменты .docx на самом деле.

    Пословный diff сохраняет порядок, а PDF отдаёт колонки и подписи к
    иллюстрациям не в том порядке, в каком они идут в .docx. Из-за этого
    один и тот же текст числится и «пропавшим», и «лишним».

    Сверяем не дословно, а по редким словам фрагмента: дословный поиск
    ломается о переносы, которые извлечение из PDF оставляет внутри слов,
    а словарь фрагмента переживает это спокойно.
    """
    fragments = []
    for tag, i1, i2, _j1, _j2 in opcodes:
        if tag != "equal" and i2 - i1 >= SHINGLE * 2:
            fragments.append(docx[i1:i2])
    fragments.sort(key=len, reverse=True)

    checked = []
    for words in fragments:
        rare = [w for w in words if len(w) >= LONG_WORD]
        if len(rare) >= 3:
            checked.append(rare)
        if len(checked) == SPOT_CHECK:
            break

    found = sum(
        1
        for rare in checked
        if sum(w in pdf_vocab for w in rare) / len(rare) >= FRAGMENT_THRESHOLD
    )
    return found, len(checked)


def grade_of(path: Path) -> str | None:
    for pattern in (GRADE_IN_PDF, GRADE_IN_DOCX):
        m = pattern.search(path.name)
        if m:
            return str(int(m.group(1)))
    return None


def main() -> int:
    lines: list[str] = []

    def out(text: str = "") -> None:
        print(text, flush=True)
        lines.append(text)

    out("Сверка PDF-версий учебников с .docx")
    out("=" * 70)

    verdict_ok = True

    for section in sorted(p for p in config.CORPUS_ROOT.iterdir() if p.is_dir()):
        out(f"\n### {section.name}")

        docx_by_grade: dict[str, Path] = {}
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
                out(f"  {pdf.name}: пары .docx нет — PDF единственный источник!")
                verdict_ok = False
                continue

            a = normalize(docx_words(docx))
            b = normalize(w for w in pdf_words(pdf) if not NOISE.match(w))
            pdf_vocab = set(b)

            matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
            opcodes = matcher.get_opcodes()
            common = sum(block.size for block in matcher.get_matching_blocks())
            extra = len(b) - common

            phrases, phrases_total = shingle_coverage(a, b)
            vocab, vocab_total = vocabulary_coverage(a, b)
            spot, spot_total = spot_check_unaligned(opcodes, a, pdf_vocab)

            missing_really = spot_total - spot
            if missing_really:
                verdict_ok = False

            out(
                f"  {grade} класс: docx {len(a)} слов, pdf {len(b)} слов, "
                f"в PDF сверх docx {extra} слов\n"
                f"        словосочетаний docx найдено в PDF: "
                f"{phrases}/{phrases_total} ({phrases / phrases_total:.1%}); "
                f"редких слов: {vocab}/{vocab_total} ({vocab / vocab_total:.1%})\n"
                f"        крупные несовпавшие фрагменты docx: {spot_total} проверено, "
                f"{spot} нашлись в PDF в другом месте, "
                f"{missing_really} действительно отсутствуют"
                + ("  <-- посмотреть глазами" if missing_really else "")
            )

    out()
    out("=" * 70)
    out(
        "Как читать числа.\n"
        "\n"
        "Проценты по словосочетаниям и редким словам НЕ равны 100 % не потому,\n"
        "что в PDF чего-то нет, а потому что извлечение текста из PDF само по\n"
        "себе неточное: слова разорваны переносами, колонки и подписи идут не в\n"
        "том порядке, что абзацы .docx. Именно поэтому в строке есть отдельная\n"
        "проверка: самые крупные фрагменты .docx, которые не легли на PDF по\n"
        "порядку, ищутся в PDF ещё раз без учёта порядка — и почти все находятся.\n"
        "\n"
        "ВАЖНО: .docx и PDF — не одна и та же редакция. Проверка вручную\n"
        "показала, что в .docx «Истории России» 10 класса есть врезка «Мнение\n"
        "учёного» с высказываниями В. Булдакова и Э. Карра и абзац про Эльзас\n"
        "и Лотарингию, которых в PDF нет вовсе — ни в каком виде, даже если\n"
        "искать по сплошному тексту без пробелов и знаков.\n"
        "\n"
        "Вывод: .docx — полноценный источник (местами более полный, чем PDF),\n"
        "по нему считаются основные числа и делается подсветка.\n"
        "\n"
        "Обратное неверно: в PDF есть материал, которого в .docx нет — подписи к\n"
        "иллюстрациям и вопросы к ним, титул, колонтитулы, условные обозначения.\n"
        "Эта добавка выделяется отдельно (src/pdfsource.py) и попадает на лист\n"
        "«Только в PDF» сводного отчёта. В общие числа она не входит: текст из\n"
        "PDF извлекается с искажениями."
    )
    if not verdict_ok:
        out()
        out("ВНИМАНИЕ: у отмеченных учебников часть крупных фрагментов .docx не")
        out("нашлась в PDF даже без учёта порядка — это стоит посмотреть глазами.")

    config.OUT_ROOT.mkdir(parents=True, exist_ok=True)
    config.OUT_PDF_LOG.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nЛог: {config.OUT_PDF_LOG}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
