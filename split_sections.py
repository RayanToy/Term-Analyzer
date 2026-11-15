"""
Разбивка выделенного учебника (docx с примечаниями) на параграфы: каждый §
в отдельном docx.

Границы параграфов — по оглавлению PDF (как в отчётах: § — страницы от его
начала до начала следующего раздела; страницы введения глав, приложений и
указателей ни в один § не входят). Страницы Word и учебника совпадают не
везде (в 11 кл. docx свёрстан иначе), поэтому страница учебника для каждого
верхнего элемента docx (абзаца, таблицы) определяется по его тексту в PDF.

В файл параграфа идут его элементы, примечания к ним и только нужные ему
картинки. Если последний элемент несёт свойства раздела, они становятся
свойствами конца документа (иначе в конце появилась бы пустая страница).

Запуск:
    python split_sections.py --docx <выделенный docx> --grade 10
"""
from __future__ import annotations

import argparse
import posixpath
import re
import statistics
import zipfile
from pathlib import Path
from types import SimpleNamespace

from lxml import etree

import config
import textbook_report as tr

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
DROPPABLE = ("/image", "/hyperlink", "/oleObject")
PARSER = etree.XMLParser(huge_tree=True)


def _own_paragraphs(el) -> list[str]:
    """Тексты абзацев элемента, включая надписи, без запасных копий надписей (mc:Fallback)."""
    out = []
    for p in el.iter(W + "p"):
        if any(a.tag == MC_FALLBACK for a in p.iterancestors()):
            continue
        # текст самого абзаца (без вложенных абзацев надписей — они идут отдельно)
        t = "".join(x.text or "" for x in p.iter(W + "t")
                    if next(x.iterancestors(W + "p")) is p)
        if t.strip():
            out.append(t)
    return out


def page_blocks(body) -> list[tuple[int, int]]:
    """
    Страницы-разделы: [(первый элемент, элемент после последнего)]. Раздел,
    начинающийся «на текущей странице» (continuous), — продолжение той же
    страницы.
    """
    ends = []
    for i, el in enumerate(body):
        sp = el if el.tag == W + "sectPr" else el.find(f"{W}pPr/{W}sectPr")
        if sp is not None:
            t = sp.find(W + "type")
            ends.append((i, t is not None and t.get(W + "val") == "continuous"))
    blocks, start = [], 0
    for k, (i, _) in enumerate(ends):
        nxt_continuous = k + 1 < len(ends) and ends[k + 1][1]
        if not nxt_continuous:
            blocks.append((start, i + 1))
            start = i + 1
    if start < len(body):
        blocks.append((start, len(body)))
    return blocks


def element_pages(body, pdf: dict) -> list[int | None]:
    """Страница учебника для каждого верхнего элемента body (по тексту в PDF)."""
    owners, texts = [], []
    for i, el in enumerate(body):
        for t in _own_paragraphs(el):
            owners.append(i)
            texts.append(t)
    pm = tr.PageMap(SimpleNamespace(texts=texts, stream=None), pdf)
    found: dict[int, list[int]] = {}
    for i, t, p in zip(owners, texts, pm.para_page):
        if p is not None and len(tr.compact(t)) >= 12:
            found.setdefault(i, []).append(p)
    pages, cur = [], None
    for i in range(len(body)):
        if i in found:
            cur = int(statistics.median_low(found[i]))
        pages.append(cur)
    # В 10 кл. каждая страница — свой раздел Word (разрыв «со следующей
    # страницы»): страница раздела — по всему его тексту, а не по отдельным
    # строкам (короткая первая строка заголовка «Биология» в надписи иначе
    # уходит на предыдущую страницу).
    blocks = page_blocks(body)
    if len(blocks) > 10:
        for i0, i1 in blocks:
            got = [p for i in range(i0, i1) for p in found.get(i, [])]
            if got:
                page = int(statistics.median_low(got))
                for i in range(i0, i1):
                    pages[i] = page
            else:
                for i in range(i0, i1):
                    pages[i] = pages[i0 - 1] if i0 else None
    # Абзац из одного номера страницы N стоит в docx в начале страницы N: с него
    # страница не меньше N (заголовки из коротких строк сами по себе по PDF
    # определяются плохо — «Популяционно-» / «видовой уровень:» есть и на
    # странице главы).
    floor = 0
    for i, el in enumerate(body):
        own = _own_paragraphs(el)
        if len(own) == 1 and re.fullmatch(r"\s*\d{1,3}\s*", own[0]):
            n = int(own[0])
            if pages[i] is not None and floor < n <= pages[i] + 3 and n >= pages[i] - 1:
                floor = n
        if pages[i] is not None and pages[i] < floor:
            pages[i] = floor
    return pages


def section_ranges(pages: list[int | None], sections: list[dict]) -> list[tuple[dict, int, int]]:
    """[(параграф, первый элемент, элемент после последнего)] — по возрастанию, без пересечений."""
    out, pos, n = [], 0, len(pages)
    for s in sections:
        i0 = next((i for i in range(pos, n) if pages[i] is not None and pages[i] >= s["start"]), None)
        if i0 is None:
            continue
        i1 = next((i for i in range(i0, n) if pages[i] is not None and pages[i] > s["end"]), n)
        out.append((s, i0, i1))
        pos = i1
    return out


def _r_refs(root) -> set[str]:
    """Все r:* ссылки (r:id, r:embed, r:link …) в документе."""
    refs = set()
    for el in root.iter():
        for k, v in el.attrib.items():
            if k.startswith("{" + R_NS + "}"):
                refs.add(v)
    return refs


def write_section(src: Path, doc_xml: bytes, comments_xml: bytes | None, rels_xml: bytes,
                  i0: int, i1: int, out: Path) -> dict:
    root = etree.fromstring(doc_xml, PARSER)
    body = root.find(W + "body")
    children = list(body)
    body_sect = children[-1] if children[-1].tag == W + "sectPr" else None
    for el in children[i1:]:
        if el is not body_sect:
            body.remove(el)
    for el in children[:i0]:
        body.remove(el)
    # свойства раздела последнего элемента -> свойства конца документа
    kept = [el for el in body if el is not body_sect]
    last_sect = kept[-1].find(f"{W}pPr/{W}sectPr") if kept and kept[-1].tag == W + "p" else None
    if last_sect is not None and body_sect is not None:
        last_sect.getparent().remove(last_sect)
        body.replace(body_sect, last_sect)

    refs = {r.get(W + "id") for r in body.iter(W + "commentReference")}
    new_comments = None
    if comments_xml is not None:
        croot = etree.fromstring(comments_xml, PARSER)
        for c in list(croot):
            if c.get(W + "id") not in refs:
                croot.remove(c)
        new_comments = etree.tostring(croot, xml_declaration=True, encoding="UTF-8", standalone=True)

    used = _r_refs(root)
    rroot = etree.fromstring(rels_xml, PARSER)
    for rel in list(rroot):
        if rel.get("Type", "").endswith(DROPPABLE) and rel.get("Id") not in used:
            rroot.remove(rel)
    keep_media = {posixpath.normpath(posixpath.join("word", rel.get("Target")))
                  for rel in rroot if rel.get("TargetMode") != "External"}

    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            if name.startswith("word/media/") and name not in keep_media:
                continue
            if name == "word/document.xml":
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif name == "word/comments.xml" and new_comments is not None:
                data = new_comments
            elif name == "word/_rels/document.xml.rels":
                data = etree.tostring(rroot, xml_declaration=True, encoding="UTF-8", standalone=True)
            else:
                data = zin.read(name)
            zout.writestr(item, data)
    return {"comments": len(refs), "media": sum(1 for m in keep_media if m.startswith("word/media/"))}


def _safe(name: str, limit: int = 90) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', " ", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name if len(name) <= limit else name[:limit].rsplit(" ", 1)[0].rstrip(" ,.")


def split(docx: Path, grade: int, out_dir: Path | None = None) -> list[Path]:
    pdf = tr.load_pdf(config.SOURCE_PDF[grade])
    sections, _ = tr.parse_toc(pdf)
    with zipfile.ZipFile(docx) as z:
        doc_xml = z.read("word/document.xml")
        comments_xml = z.read("word/comments.xml") if "word/comments.xml" in z.namelist() else None
        rels_xml = z.read("word/_rels/document.xml.rels")
    body = etree.fromstring(doc_xml, PARSER).find(W + "body")
    pages = element_pages(body, pdf)
    ranges = section_ranges(pages, sections)
    if len(ranges) != len(sections):
        missing = sorted({s["num"] for s in sections} - {s["num"] for s, _, _ in ranges})
        print(f"[Параграфы] {docx.name}: не найдены в docx § {missing}")
    out_dir = out_dir or docx.parent / f"Параграфы_{grade} класс"
    files = []
    for s, i0, i1 in ranges:
        out = out_dir / f"§ {s['num']:02d}. {_safe(s['title'])}.docx"
        info = write_section(docx, doc_xml, comments_xml, rels_xml, i0, i1, out)
        files.append(out)
        print(f"[Параграфы] § {s['num']:>2} стр. {s['start']}–{s['end']}: элементы {i0}–{i1 - 1}, "
              f"примечаний {info['comments']}, картинок {info['media']} -> {out.name}")
    return files


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--docx", required=True, type=Path)
    ap.add_argument("--grade", required=True, type=int)
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()
    split(args.docx, args.grade, args.out)


if __name__ == "__main__":
    main()
