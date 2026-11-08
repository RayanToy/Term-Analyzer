"""
PDF-черновик для издателя: слева страница учебника с подсветкой терминов
и новых слов, справа рекомендации по этой странице.

Рекомендации строятся по данным отчёта (textbook_report.collect):
  • новое слово или термин с незнакомым словом без определения при первом
    упоминании (и пример определения из другого учебника, если нашёлся);
  • определение дано позже первого упоминания;
  • слово снова встречается после большого перерыва;
  • термин давно не встречался в прошлых классах;
  • высокая плотность терминов на странице и в параграфе;
  • много новых слов на одной странице.

В PDF попадают только страницы, по которым есть рекомендации.
"""
from __future__ import annotations

import statistics
import tempfile
from collections import defaultdict
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont

import config
from textbook_report import FAR_PAGES, FIG_STATUS, LONG_GAP_YEARS, compact, load_pdf, parse_toc, term_annotations

DPI = 150
CANVAS = (int(11.69 * DPI), int(8.27 * DPI))      # A4 альбомная
MARGIN = 28
FONT_DIR = Path(r"C:\Windows\Fonts")
YELLOW = (255, 226, 0, 120)
PINK = (255, 80, 170, 105)
GREEN = (40, 220, 70, 100)


def term_fill(sources: dict | None, key: str):
    """Фон термина как в Word: ФРП — розовый, указатель — зелёный, словарь — жёлтый."""
    srcs = (sources or {}).get(key, ())
    return PINK if "frp" in srcs else GREEN if "index" in srcs else YELLOW
TURQUOISE = (0, 205, 215, 110)
RED = (220, 30, 30, 255)
PURPLE = (120, 60, 170)
INK = (30, 30, 30)
GREY = (110, 110, 110)
PANEL_BG = (247, 247, 244)
# Страница «плотная», если терминов на 100 слов не меньше, чем на 90 % страниц
DENSE_PAGE_QUANTILE = 0.9
MIN_PAGE_WORDS = 120
MANY_NEW_ON_PAGE = 4


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT_DIR / ("arialbd.ttf" if bold else "arial.ttf")), size)


def wrap(draw: ImageDraw.ImageDraw, text: str, f, width: int) -> list[str]:
    lines = []
    for para in text.split("\n"):
        line = ""
        for word in para.split(" "):
            cand = f"{line} {word}".strip()
            if draw.textlength(cand, font=f) <= width or not line:
                line = cand
            else:
                lines.append(line)
                line = word
        lines.append(line)
    return lines


# =============================================================================
# Рекомендации
# =============================================================================

def build_recommendations(a, R) -> dict[int, list[dict]]:
    """
    {страница: [{n, kind, key, text, frame}]} — frame=True: слово на
    странице обводится красным (нет определения при первом упоминании).
    """
    recs = defaultdict(list)
    sections = R["sections"]
    sec_of_page = {}
    for s in sections:
        for p in range(s["start"], s["end"] + 1):
            sec_of_page[p] = s

    def add(page, kind, key, text, frame=False):
        if page:
            recs[page].append({"kind": kind, "key": key, "text": text, "frame": frame})

    # --- Новые слова и термины с незнакомым словом без определения ------------
    items = [("new", k, R["new_defs"][k], R["new_info"][k]) for k in R["new_occ"]]
    sources = a.get("term_sources") or {}
    is_frp = lambda k: "frp" in sources.get(k, ())
    items += [("term", k, R["term_defs"][k], R["term_info"][k]) for k in R["term_occ"]
              if k in a.new_terms or is_frp(k)]
    for kind, k, d, inf in items:
        first = d["first_page"]
        if kind == "new":
            what = f"новое слово «{k}»"
        elif is_frp(k):
            what = f"термин ФРП «{k}» (требуется рабочей программой" + (
                ", в нём незнакомое ученику слово)" if k in a.new_terms else ")")
        else:
            what = f"термин «{k}» (в нём незнакомое ученику слово)"
        sample = R["other_defs"].get(k, "")
        sample_txt = f"\nОбразец из другого учебника — {sample}" if sample else ""
        if d["status"] == "нет":
            add(first, "nodef", k, f"{what.capitalize()}: определения в учебнике нет. "
                                   f"Дать определение при первом упоминании.{sample_txt}", True)
        elif d["status"] == FIG_STATUS:
            figs = ", ".join(d["figures"][:3])
            add(first, "fig", k, f"{what.capitalize()}: определения в тексте нет, но слово показано на рисунке "
                                 f"({figs}). Проверить, понятно ли значение из рисунка; при необходимости "
                                 f"дать определение в тексте или в подписи.{sample_txt}", True)
        elif d["status"] == "выделено жирным, определения нет":
            add(first, "nodef", k, f"{what.capitalize()} выделено жирным, но не определено. "
                                   f"Добавить определение.{sample_txt}", True)
        elif d["status"] == "только пояснение":
            add(first, "explain", k, f"{what.capitalize()}: есть только пояснение "
                                     f"(с. {d['page']}). При необходимости дать полное определение.", False)
        elif d["page"] and first and d["page"] > first:
            add(first, "late", k, f"{what.capitalize()} впервые встречается здесь, а определение дано "
                                  f"только на с. {d['page']}. Перенести определение к первому упоминанию "
                                  f"или дать ссылку на с. {d['page']}.", True)

        # Повторная встреча после большого перерыва
        pages = [p for p in inf["pages"] if p]
        for x, y in zip(pages, pages[1:]):
            if y - x >= FAR_PAGES:
                add(y, "far", k, f"«{k}» снова встречается через {y - x} стр. (до этого — с. {x}). "
                                 f"Напомнить значение слова.")

    # --- Термины без определения, которые ученик может не помнить --------------
    prior = R["prior"]
    for k, m in R["term_prior"].items():
        if k in a.new_terms or is_frp(k) or R["term_defs"][k]["status"] == "есть":
            continue
        verdict = prior.verdict(m)
        if not verdict.startswith(("не встречалось", "редко", "давно", "в биологии давно")):
            continue
        bio, allc = int(sum(m["bio"].values())), int(sum(m["all"].values()))
        last = f", последний раз в {m['last']} кл." if m["last"] else ""
        add(R["term_defs"][k]["first_page"], "old", k,
            f"Термин «{k}»: определения в учебнике нет, а ученик может его не помнить — в прошлых "
            f"классах встречался в учебниках биологии {bio} раз, во всех учебниках {allc} раз{last}. "
            f"Дать определение или напомнить значение.", True)

    # --- Плотность по страницам -------------------------------------------------
    word_sec = R["word_sec"]
    page_terms = defaultdict(int)
    page_new = defaultdict(set)
    pm = R["pm"]
    for k, inf in R["term_info"].items():
        for i, p in zip(inf["idx"], inf["pages"]):
            if i < len(word_sec) and word_sec[i] and p:
                page_terms[p] += 1
    for k, inf in R["new_info"].items():
        for i, p in zip(inf["idx"], inf["pages"]):
            if p:
                page_new[p].add(k)
    dens = {p: page_terms[p] / n * 100 for p, n in R["page_words"].items() if n >= MIN_PAGE_WORDS}
    vals = sorted(dens.values())
    cut = vals[int(len(vals) * DENSE_PAGE_QUANTILE)] if vals else 0
    mean = statistics.mean(vals) if vals else 0
    R["page_density"], R["page_density_cut"], R["page_density_mean"] = dens, cut, mean
    for p, d in dens.items():
        if d >= cut:
            add(p, "dense", "", f"Высокая плотность терминов: {d:.1f} на 100 слов (в среднем по учебнику "
                                f"{mean:.1f}). Разгрузить текст: пояснить термины, добавить пример или схему.")
    for p, words in page_new.items():
        if len(words) >= MANY_NEW_ON_PAGE:
            add(p, "many", "", f"На странице {len(words)} новых слов: {', '.join(sorted(words))}. "
                               f"Проверить, все ли они нужны и объяснены.")

    # --- Плотность параграфа — на его первой странице ------------------------------
    sd = R["section_density"]
    for s in sections:
        if sd.get(s["num"], 0) >= R["density_q3"]:
            add(s["start"], "sec", "", f"§ {s['num']} — один из самых насыщенных терминами параграфов: "
                                       f"{sd[s['num']]:.1f} на 100 слов основного текста.")

    order = {"nodef": 0, "fig": 1, "late": 2, "explain": 3, "old": 4, "far": 5, "many": 6, "dense": 7, "sec": 8}
    for p in recs:
        recs[p].sort(key=lambda r: order[r["kind"]])
        for n, r in enumerate(recs[p], 1):
            r["n"] = n
    return recs


# =============================================================================
# Подсветка на странице PDF
# =============================================================================

def page_marks(a, R) -> dict[int, list[tuple[str, str, str]]]:
    """{страница: [(вид 'term'|'new', ключ, текст вхождения)]} в порядке текста."""
    marks = defaultdict(list)
    text = a.stream.text
    for kind, occ in (("new", R["new_occ"]), ("term", R["term_occ"])):
        info = R["new_info"] if kind == "new" else R["term_info"]
        for k, lst in occ.items():
            for (s, e), p in zip(lst, info[k]["pages"]):
                if p:
                    marks[p].append((s, kind, k, text[s:e]))
    return {p: [(s, kind, k, surf) for s, kind, k, surf in sorted(v)] for p, v in marks.items()}


def locate(words: list, surface: str, used: set[int]) -> list[int]:
    """Индексы слов PDF, которыми набрано вхождение (перенос «шапе-» + «ронов» тоже)."""
    targets = [compact(x) for x in surface.split() if compact(x)]
    found, start = [], 0
    for t in targets:
        hit = None
        for j in range(start, len(words)):
            if j in used:
                continue
            raw = words[j][0]
            cw = compact(raw)
            if not cw:
                continue
            if cw == t:
                hit = [j]
                break
            if raw.rstrip().endswith(("-", "‑", "\u00ad")) and len(cw) >= 3 and t.startswith(cw):
                hit = [j] + ([j + 1] if j + 1 < len(words) and t.endswith(compact(words[j + 1][0])) else [])
                break
        if hit is None:
            return found
        found += hit
        start = hit[-1] + 1
    return found


# =============================================================================
# Вёрстка
# =============================================================================

def render_page(src: pdfium.PdfDocument, pdf: dict, page_no: int, marks, recs, header: str,
                notes: dict | None = None, sources: dict | None = None) -> Image.Image:
    canvas = Image.new("RGB", CANVAS, "white")
    draw = ImageDraw.Draw(canvas)
    # --- Левая часть: страница учебника --------------------------------------
    pinfo = pdf["pages"][page_no - 1]
    target_h = CANVAS[1] - 2 * MARGIN
    scale = target_h / pinfo["h"]
    img = src[page_no - 1].render(scale=scale).to_pil().convert("RGBA")
    over = Image.new("RGBA", img.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(over)
    words = pinfo["words"]
    used: set[int] = set()
    framed = {r["key"] for r in recs if r["frame"] and r["key"]}
    numbered = {}
    for r in recs:
        if r["key"] and r["key"] not in numbered:
            numbered[r["key"]] = r["n"]
    badge_done = set()
    notes = notes or {}
    note_font = font(12, True)
    for start, kind, key, surface in marks:
        idx = locate(words, surface, used)
        used.update(idx)
        note = notes.get(start)
        if idx and note:
            # Числа упоминаний в прошлых классах, как в Word: сверху — биология, снизу — все учебники
            _, _, sup, sub = (None, *note)
            _, nx0, ntop, nx1, nbottom = words[idx[-1]]
            x = nx1 * scale + 2
            for text, y, color in ((sup, ntop * scale - 9, (30, 123, 52, 255)),
                                   (sub, nbottom * scale - 6, (31, 78, 154, 255))):
                w = od.textlength(text, font=note_font)
                od.rectangle((x - 1, y + 1, x + w + 1, y + 13), fill=(255, 255, 255, 200))
                od.text((x, y - 1), text, font=note_font, fill=color)
        for j in idx:
            _, x0, top, x1, bottom = words[j]
            box = (x0 * scale - 1, top * scale - 1, x1 * scale + 1, bottom * scale + 1)
            od.rectangle(box, fill=term_fill(sources, key) if kind == "term" else TURQUOISE)
            if key in framed and key not in badge_done:
                od.rectangle(box, outline=RED, width=3)
        if idx and key in numbered and key not in badge_done:
            _, x0, top, x1, bottom = words[idx[0]]
            cx, cy = x0 * scale - 2, top * scale - 2
            od.ellipse((cx - 13, cy - 13, cx + 13, cy + 13), fill=RED)
            label = str(numbered[key])
            f = font(16, True)
            od.text((cx - od.textlength(label, font=f) / 2, cy - 9), label, font=f, fill="white")
            badge_done.add(key)
    img = Image.alpha_composite(img, over).convert("RGB")
    canvas.paste(img, (MARGIN, MARGIN))
    draw.rectangle((MARGIN - 1, MARGIN - 1, MARGIN + img.width, MARGIN + img.height), outline=(190, 190, 190))

    # --- Правая часть: рекомендации --------------------------------------------
    x0 = MARGIN * 2 + img.width
    width = CANVAS[0] - x0 - MARGIN
    draw.rectangle((x0 - 10, MARGIN, CANVAS[0] - MARGIN + 10, CANVAS[1] - MARGIN), fill=PANEL_BG)
    for size in (21, 19, 17, 15):
        f, fb = font(size), font(size, True)
        y = MARGIN + 14
        blocks = []
        y += draw_textblock(None, x0, y, header, font(size + 4, True), width) + 14
        for r in recs:
            lines = wrap(draw, r["text"], f, width - 40)
            blocks.append((r, lines))
            y += len(lines) * (size + 6) + 12
        if y <= CANVAS[1] - MARGIN - 90:
            break
    y = MARGIN + 14
    y += draw_textblock(draw, x0, y, header, font(size + 4, True), width) + 14
    limit = CANVAS[1] - MARGIN - 90
    if not blocks:
        draw.text((x0, y), "Замечаний нет.", font=font(21), fill=GREY)
    for i, (r, lines) in enumerate(blocks):
        h = len(lines) * (size + 6)
        if y + h > limit:
            draw.text((x0, y), f"… ещё замечаний: {len(blocks) - i}", font=fb, fill=GREY)
            break
        color = (RED if r["kind"] in ("nodef", "late") else PURPLE if r["kind"] == "fig"
                 else (0, 140, 150) if r["kind"] in ("explain", "far", "old", "many") else (190, 120, 0))
        draw.ellipse((x0, y + 1, x0 + 26, y + 27), fill=color)
        lab = str(r["n"])
        draw.text((x0 + 13 - draw.textlength(lab, font=font(16, True)) / 2, y + 4), lab, font=font(16, True), fill="white")
        for ln in lines:
            draw.text((x0 + 38, y), ln, font=f, fill=INK)
            y += size + 6
        y += 12
    legend = ("Термины: розовым — ФРП, зелёным — указатель учебника, жёлтым — словарь терминов; бирюзовым — "
              "новые слова; красная рамка и номер — к слову есть замечание.\n"
              "Числа у термина — упоминания по классам («5:12» — в 5 классе 12 раз): верхние (зелёные) — "
              "учебники биологии, нижние (синие) — все учебники прошлых классов.")
    gl = wrap(draw, legend, font(14), width)
    for i, ln in enumerate(gl):
        draw.text((x0, CANVAS[1] - MARGIN - 12 - (len(gl) - i) * 18), ln, font=font(14), fill=GREY)
    return canvas


def draw_textblock(draw, x, y, text, f, width) -> int:
    tmp = draw or ImageDraw.Draw(Image.new("RGB", (10, 10)))
    lines = wrap(tmp, text, f, width)
    if draw:
        for i, ln in enumerate(lines):
            draw.text((x, y + i * (f.size + 6)), ln, font=f, fill=INK)
    return len(lines) * (f.size + 6)


def cover_page(a, R, recs, n_pages_total: int) -> Image.Image:
    img = Image.new("RGB", CANVAS, "white")
    d = ImageDraw.Draw(img)
    x, y, w = 90, 80, CANVAS[0] - 180
    y += draw_textblock(d, x, y, f"Рекомендации издателю — черновик", font(40, True), w) + 10
    y += draw_textblock(d, x, y, f"Биология, {a.grade} класс (Пасечник, Каменский). Базовый уровень", font(26), w) + 30
    kinds = defaultdict(int)
    for lst in recs.values():
        for r in lst:
            kinds[r["kind"]] += 1
    lines = [
        "Слева — страница учебника. Термины: розовым — из ФРП (рабочей программы), зелёным — из указателя "
        "терминов учебника, жёлтым — из общего словаря терминов; бирюзовым — новые для ученика слова "
        f"(в прошлых классах встречались меньше {a.min_freq} раз). Красной рамкой и номером отмечены слова, "
        "к которым относится замечание справа.",
        "",
        f"В черновике все {n_pages_total} страниц учебника; замечания есть на {len(recs)} из них.",
        "Числа у термина (при первом упоминании и после перерыва в 20 и больше страниц) — упоминания в прошлых "
        "классах по классам, «5:12» — в 5 классе 12 раз: верхняя строка, зелёная — учебники биологии, нижняя, "
        "синяя — все школьные учебники. Классы без упоминаний не показаны, «0» — ни разу, «т» — тысяч.",
        "",
        "Виды замечаний:",
        f"  • нет определения при первом упоминании нового слова или термина — {kinds['nodef']}",
        f"  • определения нет, но слово показано на рисунке (отдельный случай) — {kinds['fig']}",
        f"  • определение дано позже первого упоминания — {kinds['late']}",
        f"  • есть только пояснение (синоним, «т. е.», происхождение слова) — {kinds['explain']}",
        f"  • термин без определения, который ученик может не помнить (не встречался, встречался редко "
        f"или {LONG_GAP_YEARS}+ лет назад) — {kinds['old']}",
        f"  • слово снова встречается после перерыва в {FAR_PAGES}+ страниц — {kinds['far']}",
        f"  • много новых слов на одной странице ({MANY_NEW_ON_PAGE}+) — {kinds['many']}",
        f"  • высокая плотность терминов на странице (10 % самых плотных) — {kinds['dense']}",
        f"  • параграф с высокой плотностью терминов (25 % самых плотных) — {kinds['sec']}",
        "",
        "Определения и пояснения найдены автоматически по шаблонам («X — это…», «…называют X» и т. п.), "
        "поэтому каждое замечание стоит проверить. Это черновик для обсуждения вида отчёта.",
    ]
    for ln in lines:
        y += draw_textblock(d, x, y, ln, font(24), w) + 4
    return img


def write_pdf_report(a, R, out_path: Path, pdf_path: Path | None = None) -> Path:
    pdf_path = Path(pdf_path or config.SOURCE_PDF[a.grade])
    pdf = load_pdf(pdf_path)
    recs = build_recommendations(a, R)
    marks = page_marks(a, R)
    notes = R.get("notes") or term_annotations(a)["by_start"]
    sec_of_page = {}
    for s in R["sections"]:
        for p in range(s["start"], s["end"] + 1):
            sec_of_page[p] = s
    src = pdfium.PdfDocument(str(pdf_path))
    out = pdfium.PdfDocument.new()
    w_pt, h_pt = 842, 595
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        pages = [cover_page(a, R, recs, len(pdf["pages"]))]
        files = []
        for i, img in enumerate(pages):
            fp = Path(tmp) / f"cover{i}.jpg"
            img.save(fp, "JPEG", quality=85)
            files.append(fp)
        for p in range(1, len(pdf["pages"]) + 1):
            s = sec_of_page.get(p)
            header = f"Страница {p}" + (f" · § {s['num']}. {s['title']}" if s else "")
            img = render_page(src, pdf, p, marks.get(p, []), recs.get(p, []), header, notes,
                              a.get("term_sources"))
            fp = Path(tmp) / f"p{p:04d}.jpg"
            img.save(fp, "JPEG", quality=70)
            files.append(fp)
        for fp in files:
            page = out.new_page(w_pt, h_pt)
            image = pdfium.PdfImage.new(out)
            image.load_jpeg(str(fp), inline=False)
            image.set_matrix(pdfium.PdfMatrix().scale(w_pt, h_pt))
            page.insert_obj(image)
            page.gen_content()
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out.save(str(out_path))
        out.close()
    src.close()
    print(f"[PDF-отчёт] {out_path}: страниц {len(pdf['pages'])}, с замечаниями {len(recs)}")
    return out_path
