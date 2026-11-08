"""
Финальный шаг: PDF размеченных учебников через Word и исправление пометок,
которые Word не смог разместить.

Числа упоминаний у терминов стоят в поле EQ (две строки друг над другом).
Если вёрстка сконвертированного учебника прижимает строку к картинке так,
что в ней остаётся место на пару букв, поле не помещается и Word печатает
«Ошибка!». Заранее это видно только самому Word, поэтому:
  1. документ экспортируется в PDF;
  2. по тексту PDF находятся страницы с «Ошибка!»;
  3. на этих страницах поля-пометки заменяются обычным надстрочным
     и подстрочным текстом тех же цветов;
  4. документ сохраняется и экспортируется снова (до 4 раз).

Word запускается отдельным процессом; открытые пользователем окна Word
не затрагиваются.

Запуск: python word_finalize.py <docx> [<docx> ...]
"""
from __future__ import annotations

import re
import sys
import time
from pathlib import Path

import pypdfium2 as pdfium

ERROR_MARK = "Ошибка!"
EQ_RE = re.compile(r"\\up\d+\((.*?)\);\\s\\do\d+\((.*?)\)\)")
GREEN = 0x1E + 0x7B * 256 + 0x34 * 65536     # 1E7B34 в формате BGR для Word
BLUE = 0x1F + 0x4E * 256 + 0x9A * 65536      # 1F4E9A
WD_FORMAT_PDF = 17
WD_ACTIVE_END_PAGE = 3


def word_pids() -> set[int]:
    import subprocess
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True).stdout.decode("cp866", errors="ignore")
    return {int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"WINWORD')}


def error_pages(pdf_path: Path) -> list[int]:
    doc = pdfium.PdfDocument(str(pdf_path))
    pages = [i + 1 for i in range(len(doc)) if ERROR_MARK in doc[i].get_textpage().get_text_bounded()]
    doc.close()
    return pages


def flatten_fields(doc, pages: set[int]) -> int:
    """Поля-пометки на указанных страницах — в обычный надстрочный/подстрочный текст."""
    fixed = 0
    for i in range(doc.Fields.Count, 0, -1):
        f = doc.Fields(i)
        code = f.Code.Text
        if "\\o\\al" not in code:
            continue
        m = EQ_RE.search(code)
        # Страница результата поля совпадает со страницей PDF; страница кода поля —
        # нет (Word считает её так, будто коды полей показаны)
        if not m or f.Result.Information(WD_ACTIVE_END_PAGE) not in pages:
            continue
        sup, sub = m.group(1), m.group(2)
        pos = f.Code.Start - 1
        f.Delete()
        rng = doc.Range(pos, pos)
        rng.InsertAfter(sup + sub)
        for start, text, color, is_sup in ((pos, sup, GREEN, True), (pos + len(sup), sub, BLUE, False)):
            r = doc.Range(start, start + len(text))
            r.Font.Size = 6
            r.Font.Bold = True
            r.HighlightColorIndex = 0
            r.Font.Color = color
            r.Font.Superscript = is_sup
            r.Font.Subscript = not is_sup
        fixed += 1
    return fixed


def finalize(word, docx: Path, rounds: int = 4) -> None:
    pdf = docx.with_suffix(".pdf")
    doc = word.Documents.Open(str(docx), False, False)
    try:
        for attempt in range(rounds):
            doc.ExportAsFixedFormat(str(pdf), WD_FORMAT_PDF)
            bad = error_pages(pdf)
            if not bad:
                print(f"[Word] {pdf.name}: готово, ошибок нет")
                break
            fixed = flatten_fields(doc, set(bad))
            print(f"[Word] {docx.name}: «{ERROR_MARK}» на стр. {bad}; пометок переведено в обычный текст: {fixed}")
            doc.Save()
            if not fixed:
                print("[Word] пометок на этих страницах нет — ошибки от других полей, оставляю как есть")
                break
        else:
            print(f"[Word] {pdf.name}: после {rounds} попыток остались ошибки на стр. {error_pages(pdf)}")
    finally:
        doc.Close(0)


def main(paths: list[str]) -> None:
    import win32com.client
    before = word_pids()
    word = win32com.client.DispatchEx("Word.Application")
    time.sleep(1)
    mine = word_pids() - before
    if not mine:
        word.Quit()
        raise SystemExit("Не удалось запустить отдельный процесс Word")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for p in paths:
            finalize(word, Path(p).resolve())
    finally:
        word.Quit()
        time.sleep(2)
        import subprocess
        for pid in mine & word_pids():
            subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


if __name__ == "__main__":
    main(sys.argv[1:])
