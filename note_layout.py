"""
Раскладка чисел у терминов без наложений.

Числа стоят в плавающих надписях (не сдвигают текст), привязанных к
первой букве термина: верхний ряд — над строкой, нижний — под ней. Если
термины стоят рядом, их надписи наезжают друг на друга; нижний ряд одной
строки может наехать и на верхний ряд следующей (они в одном межстрочном
промежутке).

Где окажется надпись, знает только Word, поэтому:
  1. Word открывает документ, сообщает страницу и примерное место каждой
     надписи (у надписи в «замещающем тексте» метка tn:<номер>:<t|b>)
     и экспортирует черновой PDF;
  2. в PDF находятся настоящие прямоугольники рядов чисел (по размеру и
     цвету шрифта) и сопоставляются с надписями: страница, текст, ряд
     (зелёный верхний / синий нижний), ближайшее место. Координаты Word
     на страницах с колонками бывают сдвинуты, а PDF показывает, как есть;
  3. на каждой странице ряды раскладываются слева направо: если ряд
     пересекается с уже поставленным, он сдвигается вправо за него
     (у правого края страницы — влево);
  4. новые сдвиги (margin-left) записываются в docx.
Текст учебника при этом не меняется: надписи в вёрстке места не занимают.
"""
from __future__ import annotations

import re
import subprocess
from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

NOTE_FONT = Path(r"C:\Windows\Fonts\arialbd.ttf")
GAP_PT = 2.0            # зазор между соседними рядами, пт
EDGE_PT = 4.0           # от правого края страницы, пт
MAX_NOTE_SIZE = 5.6     # ряды чисел набраны мельче всего остального текста
ALT_RE = re.compile(r"^tn:(\d+):([tb])$")
V_RECT = "{urn:schemas-microsoft-com:vml}rect"

PS_DUMP = r"""
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
Start-Sleep -Milliseconds 800
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Where-Object { $before -notcontains $_ })
if ($mine.Count -eq 0) { throw 'Не удалось запустить отдельный процесс Word' }
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $d = $word.Documents.Open($args[0], $false, $true)
    $d.Repaginate()
    $out = New-Object System.Text.StringBuilder
    foreach ($s in $d.Shapes) {
        $alt = $s.AlternativeText
        if ($alt -notlike 'tn:*') { continue }
        $a = $s.Anchor
        [void]$out.AppendLine(('{0}|{1}|{2}|{3}' -f $alt, $a.Information(3), $a.Information(5), $a.Information(6)))
    }
    [System.IO.File]::WriteAllText($args[1], $out.ToString())
    $d.ExportAsFixedFormat($args[2], 17)
    $d.Close(0)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Start-Sleep -Seconds 2
    foreach ($id in $mine) { Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force }
}
"""


@lru_cache(maxsize=4)
def _font(size_pt: float):
    # шрифт в 10 раз крупнее — точнее ширина; результат делится на 10
    return ImageFont.truetype(str(NOTE_FONT), int(size_pt * 10))


def text_width(text: str, size_pt: float) -> float:
    """Ширина строки жирным Arial, пт."""
    return _font(size_pt).getlength(text) / 10


def _squash(text: str) -> str:
    return re.sub(r"\s+", "", text)


# =============================================================================
# Word: страницы надписей и черновой PDF
# =============================================================================

def word_dump(docx: Path, work_dir: Path) -> tuple[list[dict], Path]:
    """Страница и примерное место каждой надписи с числами + черновой PDF."""
    txt = work_dir / "notes.txt"
    pdf = work_dir / "draft.pdf"
    script = "& {" + PS_DUMP + "} " + " ".join(
        "'" + str(p).replace("'", "''") + "'" for p in (docx, txt, pdf))
    res = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                         capture_output=True, text=True, encoding="cp866", errors="ignore")
    if res.returncode or not txt.exists() or not pdf.exists():
        raise RuntimeError(f"Word не отдал позиции надписей: {res.stderr.strip()}")
    items = []
    for line in txt.read_text(encoding="utf-8-sig").splitlines():
        parts = line.strip().split("|")
        m = ALT_RE.match(parts[0])
        if not m or len(parts) < 4:
            continue
        num = [float(x.replace(",", ".")) for x in parts[1:4]]
        items.append({"key": (int(m.group(1)), m.group(2)), "page": int(num[0]), "x": num[1], "y": num[2]})
    return items, pdf


# =============================================================================
# PDF: настоящие прямоугольники рядов чисел
# =============================================================================

def pdf_rows(pdf_path: Path) -> tuple[dict[int, list[dict]], dict[int, float]]:
    """{страница: [ряд {text, row t|b, x0, x1, top, bottom}]}, {страница: ширина}."""
    import pdfplumber
    rows_by_page, widths = {}, {}
    with pdfplumber.open(str(pdf_path)) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            widths[pno] = float(page.width)
            # Порядок символов — как в PDF: надпись рисуется целиком, поэтому её символы
            # идут подряд, даже если ряд наезжает на соседний (сортировка по x их перемешала бы)
            chars = [c for c in page.chars if c["size"] < MAX_NOTE_SIZE and c["text"].strip()]
            rows: list[dict] = []
            for c in chars:
                col = c.get("non_stroking_color") or (0, 0, 0)
                if not isinstance(col, (tuple, list)) or len(col) < 3:
                    continue
                kind = "t" if col[1] > col[2] else "b"      # зелёный — верхний, синий — нижний
                r = rows[-1] if rows else None
                if (r and r["row"] == kind and abs(r["top"] - c["top"]) < 0.8
                        and -0.5 < c["x0"] - r["x1"] < 2.2):
                    r["x1"] = c["x1"]
                    r["bottom"] = max(r["bottom"], c["bottom"])
                    r["text"] += c["text"]
                else:
                    rows.append({"text": c["text"], "row": kind, "x0": c["x0"], "x1": c["x1"],
                                 "top": c["top"], "bottom": c["bottom"]})
            rows_by_page[pno] = rows
    return rows_by_page, widths


def match_rows(items: list[dict], texts: dict, rows_by_page: dict) -> dict:
    """Надпись -> её ряд в PDF: та же страница, тот же текст и ряд, ближайшее место."""
    matched = {}
    by_page: dict[int, list[dict]] = {}
    for it in items:
        by_page.setdefault(it["page"], []).append(it)
    for page, page_items in by_page.items():
        free = list(rows_by_page.get(page, []))
        cand = []
        for it in page_items:
            text = _squash(texts.get(it["key"], ""))
            for j, r in enumerate(free):
                if r["row"] == it["key"][1] and _squash(r["text"]) == text:
                    # x у Word на страницах с колонками бывает сдвинут, y надёжен
                    dist = abs(r["top"] - it["y"]) + 0.2 * abs(r["x0"] - it["x"])
                    cand.append((dist, it["key"], j))
        used_rows, used_keys = set(), set()
        for dist, key, j in sorted(cand):
            if key in used_keys or j in used_rows:
                continue
            matched[key] = free[j]
            used_keys.add(key)
            used_rows.add(j)
    return matched


# =============================================================================
# Раскладка
# =============================================================================

def resolve(matched: dict, page_of: dict, page_w: dict) -> dict:
    """Сдвиг каждого ряда по горизонтали (пт), чтобы ряды не пересекались; только ненулевые."""
    by_page: dict[int, list] = {}
    for key, r in matched.items():
        by_page.setdefault(page_of[key], []).append((key, r))
    shifts = {}
    for page, lst in by_page.items():
        placed: list[tuple[float, float, float, float]] = []
        right = page_w.get(page, 595.0) - EDGE_PT
        for key, r in sorted(lst, key=lambda kr: (round(kr[1]["top"]), kr[1]["x0"])):
            w = r["x1"] - r["x0"]
            y0, y1 = r["top"] + 0.3, r["bottom"] - 0.3
            x0 = r["x0"]

            def hits(x):
                return [p for p in placed if p[0] < x + w + GAP_PT and x < p[1] + GAP_PT
                        and p[2] < y1 and y0 < p[3]]

            x = x0
            for _ in range(60):
                h = hits(x)
                if not h:
                    break
                x = max(p[1] for p in h) + GAP_PT
            if x + w > right:
                # у правого края — левее тех, с кем ряд пересекается
                x = min(x0, right - w)
                for _ in range(60):
                    h = hits(x)
                    if not h:
                        break
                    x = min(p[0] for p in h) - GAP_PT - w
            placed.append((x, x + w, y0, y1))
            if abs(x - x0) > 0.05:
                shifts[key] = x - x0
    return shifts


def read_notes(docx: Path) -> tuple[dict, dict]:
    """Текст и текущий margin-left каждой надписи в docx."""
    from docx import Document
    from docx.oxml.ns import qn
    doc = Document(str(docx))
    texts, lefts = {}, {}
    for el in doc.element.body.iter(V_RECT):
        m = ALT_RE.match(el.get("alt") or "")
        if not m:
            continue
        key = (int(m.group(1)), m.group(2))
        texts[key] = "".join(t.text or "" for t in el.iter(qn("w:t")))
        ml = re.search(r"margin-left:(-?[\d.]+)pt", el.get("style") or "")
        lefts[key] = float(ml.group(1)) if ml else 0.0
    return texts, lefts


def apply_shifts(docx: Path, shifts: dict, lefts: dict) -> int:
    """Записывает новые margin-left надписей в docx."""
    from docx import Document
    doc = Document(str(docx))
    done = 0
    for el in doc.element.body.iter(V_RECT):
        m = ALT_RE.match(el.get("alt") or "")
        if not m:
            continue
        key = (int(m.group(1)), m.group(2))
        if key in shifts:
            new = round(lefts.get(key, 0.0) + shifts[key], 1)
            el.set("style", re.sub(r"margin-left:-?[\d.]+pt", f"margin-left:{new}pt", el.get("style")))
            done += 1
    doc.save(str(docx))
    return done


def layout(docx: Path, work_dir: Path) -> None:
    """Раздвигает наезжающие ряды чисел у соседних терминов."""
    texts, lefts = read_notes(docx)
    if not texts:
        return
    items, draft = word_dump(docx, work_dir)
    rows, page_w = pdf_rows(draft)
    matched = match_rows(items, texts, rows)
    page_of = {it["key"]: it["page"] for it in items}
    shifts = resolve(matched, page_of, page_w)
    apply_shifts(docx, shifts, lefts)
    print(f"[Числа] {docx.name}: надписей {len(texts)}, найдено в PDF {len(matched)}, сдвинуто {len(shifts)}")
