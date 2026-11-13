"""
Одно примечание с числами на термин на каждой странице Word; примечания
страницы собраны не больше чем в MAX_BALLOONS выносок (подряд идущие термины,
по строке на термин), иначе Word не умещает выноски и сворачивает их.

Примечания ставятся у каждого вхождения термина, а на какой странице
окажется вхождение, знает только Word. Поэтому Word сообщает страницу
каждого примечания, и на каждой странице для каждого термина остаётся
только первое примечание — остальные (повторы термина на той же странице)
удаляются из docx. Текст примечания начинается с термина («мутация: био …»),
по нему повторы и определяются.
"""
from __future__ import annotations

import bisect
import copy
import math
import re
import subprocess
from pathlib import Path

MAX_BALLOONS = 8      # выносок на странице: больше Word не умещает и сворачивает их
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

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
    # Information(3) у каждого примечания — около 0,2 с (часы на учебник), поэтому
    # берутся начала страниц и позиция примечания; каждое 40-е сверяется напрямую
    $pages = $d.ComputeStatistics(2)
    for ($p = 1; $p -le $pages; $p++) {
        [void]$out.AppendLine(('P|{0}|{1}' -f $p, $d.GoTo(1, 1, $p).Start))
    }
    $n = 0
    foreach ($c in $d.Comments) {
        $chk = 0
        if ($n % 40 -eq 0) { $chk = $c.Scope.Information(3) }
        [void]$out.AppendLine(('C|{0}|{1}|{2}' -f $c.Scope.Start, $chk, ($c.Range.Text -replace '[\r\n\a]', ' ')))
        $n++
    }
    [System.IO.File]::WriteAllText($args[1], $out.ToString())
    $d.Close(0)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Start-Sleep -Seconds 2
    foreach ($id in $mine) { Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force }
}
"""


def _sq(t: str) -> str:
    return re.sub(r"\s+", "", t)


def dedupe(docx: Path, work_dir: Path) -> None:
    from docx import Document
    doc = Document(str(docx))
    if next(doc.element.body.iter(W + "commentReference"), None) is None:
        return
    comments_el = doc.part._comments_part.element
    by_id = {c.get(W + "id"): c for c in comments_el.iter(W + "comment")}
    if not by_id:
        return
    text_of = {i: "".join(t.text or "" for t in c.iter(W + "t")) for i, c in by_id.items()}
    # порядок примечаний в документе — как у Word (Comments идёт по порядку ссылок в тексте)
    order = [r.get(W + "id") for r in doc.element.body.iter(W + "commentReference")]

    txt = work_dir / "comments.txt"
    script = "& {" + PS_DUMP + "} " + " ".join("'" + str(p).replace("'", "''") + "'" for p in (docx, txt))
    res = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                         capture_output=True, text=True, encoding="cp866", errors="ignore")
    if res.returncode or not txt.exists():
        raise RuntimeError(f"Word не отдал страницы примечаний: {res.stderr.strip()}")
    starts, rows, wrong = [], [], 0
    for ln in txt.read_text(encoding="utf-8-sig").splitlines():
        parts = ln.split("|", 3)
        if parts[0] == "P" and len(parts) >= 3:
            starts.append(int(parts[2]))
        elif parts[0] == "C" and len(parts) == 4:
            page = bisect.bisect_right(starts, int(parts[1]))
            wrong += parts[2] != "0" and int(parts[2]) != page
            rows.append((page, parts[3]))
    if starts != sorted(starts) or wrong:
        print(f"[Примечания] {docx.name}: страницы примечаний определены ненадёжно "
              f"(расхождений {wrong}) — повторы не убраны")
        return
    if len(rows) != len(order) or any(_sq(text_of[i]) != _sq(r[1]) for i, r in zip(order, rows)):
        print(f"[Примечания] {docx.name}: порядок примечаний Word и docx не совпал "
              f"({len(rows)} и {len(order)}) — повторы не убраны")
        return
    seen, drop = set(), set()
    para_of = {}
    for ref in doc.element.body.iter(W + "commentReference"):
        p = ref
        while p is not None and p.tag != W + "p":
            p = p.getparent()
        para_of[ref.get(W + "id")] = p
    groups: dict = {}
    for cid, (page, _) in zip(order, rows):
        key = (page, text_of[cid].split(":", 1)[0])
        if key in seen:
            drop.add(cid)
        else:
            groups.setdefault(page, []).append(cid)
        seen.add(key)
    # Выноски одного абзаца на одной странице — в одну (по строке на термин): иначе на
    # плотной странице их десятки, Word не умещает их по высоте и сворачивает в одну строку
    merged = 0
    chunks = []
    for ids in groups.values():
        # на странице не больше MAX_BALLOONS выносок: подряд идущие термины в одну,
        # разрыв — на границе абзаца, если выноска уже набрала свою долю терминов
        target = max(1, math.ceil(len(ids) / MAX_BALLOONS))
        cur = []
        for cid in ids:
            if cur and (len(cur) >= 2 * target
                        or (len(cur) >= target and para_of[cid] is not para_of[cur[-1]])):
                chunks.append(cur)
                cur = []
            cur.append(cid)
        if cur:
            chunks.append(cur)
    for ids in chunks:
        keep = by_id[ids[0]]
        for cid in ids[1:]:
            for cp in by_id[cid].findall(W + "p"):
                np_ = copy.deepcopy(cp)
                for r in list(np_.iter(W + "r")):
                    if r.find(W + "annotationRef") is not None:
                        r.getparent().remove(r)
                keep.append(np_)
            drop.add(cid)
            merged += 1
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(doc.element.body.iter(W + tag)):
            if el.get(W + "id") in drop:
                el.getparent().remove(el)
    for ref in list(doc.element.body.iter(W + "commentReference")):
        if ref.get(W + "id") in drop:
            run = ref.getparent()
            run.getparent().remove(run)
    for cid in drop:
        comments_el.remove(by_id[cid])
    doc.save(str(docx))
    print(f"[Примечания] {docx.name}: было {len(order)}, повторов на странице убрано {len(drop) - merged}, "
          f"объединено в выноски {merged}, выносок осталось {len(order) - len(drop)}")
