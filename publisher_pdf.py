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
from textbook_report import EVERYDAY_PRIMARY, EVERYDAY_TOTAL, FAR_PAGES, FIG_STATUS, LONG_GAP_YEARS, compact, load_pdf, parse_toc, term_annotations

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
# «Определение дано позже первого упоминания» — если между ними не меньше стольких страниц
LATE_MIN_PAGES = 3


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

    everyday = R.get("everyday", set())

    def add(page, kind, key, text, frame=False):
        # Общеупотребительные термины («человек», «жизнь») знакомы всем — без замечаний
        if key and key in everyday:
            return
        if page:
            recs[page].append({"kind": kind, "key": key, "text": text, "frame": frame})

    # --- Новые слова и термины с незнакомым словом без определения ------------
    items = [("new", k, R["new_defs"][k], R["new_info"][k]) for k in R["new_occ"]]
    sources = a.get("term_sources") or {}
    is_frp = lambda k: "frp" in sources.get(k, ())
    items += [("term", k, R["term_defs"][k], R["term_info"][k]) for k in R["term_occ"]
              if k in a.new_terms or is_frp(k)]
    word_sec = R["word_sec"]

    def main_first(kind, k):
        """Первое упоминание в основном тексте параграфа (не в «Вы узнаете», вопросах, ключевых словах)."""
        inf = R["new_info" if kind == "new" else "term_info"].get(k)
        if not inf:
            return None
        for i, p in zip(inf["idx"], inf["pages"]):
            if p and i < len(word_sec) and word_sec[i]:
                return p
        return None

    def may_forget(k):
        m = R["term_prior"].get(k)
        return bool(m) and not R["prior"].verdict(m).startswith("встречалось недавно")

    for kind, k, d, inf in items:
        # Новое слово, которое в другой форме выделено как термин с определением, — без замечания
        if kind == "new" and R["term_defs"].get(k, {}).get("status") == "есть":
            continue
        first = d["first_page"]
        if kind == "new":
            what = f"новое слово «{k}»"
        elif is_frp(k):
            what = f"термин ФРП «{k}»" + (" (в нём незнакомое ученику слово)" if k in a.new_terms else "")
        else:
            what = f"термин «{k}» (в нём незнакомое ученику слово)"
        sample = R["other_defs"].get(k, "")
        sample_txt = f"\nОбразец из другого учебника — {sample}" if sample else ""
        if d["status"] == "нет":
            add(first, "nodef", k, f"{(what[:1].upper() + what[1:])}: определения в учебнике нет. "
                                   f"Дать определение при первом упоминании.{sample_txt}", True)
        elif d["status"] == FIG_STATUS:
            figs = ", ".join(d["figures"][:3])
            add(first, "fig", k, f"{(what[:1].upper() + what[1:])}: определения в тексте нет, но слово показано на рисунке "
                                 f"({figs}). Проверить, понятно ли значение из рисунка; при необходимости "
                                 f"дать определение в тексте или в подписи.{sample_txt}", True)
        elif d["status"] == "выделено жирным, определения нет":
            add(first, "nodef", k, f"{(what[:1].upper() + what[1:])} выделено жирным, но не определено. "
                                   f"Добавить определение.{sample_txt}", True)
        elif d["status"] == "только пояснение":
            add(first, "explain", k, f"{(what[:1].upper() + what[1:])}: есть только пояснение "
                                     f"(с. {d['page']}). При необходимости дать полное определение.", False)
        elif d["page"] and main_first(kind, k) and d["page"] - main_first(kind, k) >= LATE_MIN_PAGES                 and (kind == "new" or k in a.new_terms or may_forget(k))                 and not (d.get("explain_page") and d["explain_page"] <= main_first(kind, k)):
            # (если при первом упоминании уже есть пояснение — «белковой оболочки (капсида)» — не требуем переноса)
            first = main_first(kind, k)
            add(first, "late", k, f"{(what[:1].upper() + what[1:])} впервые встречается здесь, а определение дано "
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
        last = f", последний раз в {m['last']} классе" if m["last"] else ""
        add(R["term_defs"][k]["first_page"], "old", k,
            f"Термин «{k}»: определения в учебнике нет, а ученик может его не помнить — в прошлых "
            f"классах встречался в учебниках биологии {bio} раз, во всех учебниках {allc} раз{last}. "
            f"Дать определение или напомнить значение.", True)

    # --- Плотность по страницам -------------------------------------------------
    word_sec = R["word_sec"]
    page_terms = defaultdict(int)
    page_new = defaultdict(set)
    pm = R["pm"]
    for k, i, p in R["density_occ"]:
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
    R["page_terms"] = page_terms
    q = lambda f: vals[min(len(vals) - 1, int(len(vals) * f))] if vals else 0
    R["density_zones"] = (q(0.25), q(0.75), cut)
    for p, d in dens.items():
        if d >= cut:
            add(p, "dense", "", f"Высокая плотность терминов: {d:.1f} на 100 слов ({page_terms[p]} / "
                                f"{R['page_words'][p]} × 100; в среднем по учебнику {mean:.1f}). Разгрузить "
                                f"текст: пояснить термины, добавить пример или схему.")
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
                notes: dict | None = None, sources: dict | None = None,
                density: dict | None = None) -> Image.Image:
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
        if y <= CANVAS[1] - MARGIN - 90 - DENSITY_BLOCK_H:
            break
    y = MARGIN + 14
    y += draw_textblock(draw, x0, y, header, font(size + 4, True), width) + 14
    limit = CANVAS[1] - MARGIN - 90 - DENSITY_BLOCK_H
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
    if density is not None:
        draw_density(draw, x0, CANVAS[1] - MARGIN - 90 - DENSITY_BLOCK_H + 8, width, density)
    legend = ("Термины: розовым — ФРП, зелёным — указатель учебника, жёлтым — словарь терминов; бирюзовым — "
              "новые слова; красная рамка и номер — к слову есть замечание.\n"
              "Числа у термина — упоминания по классам («5:12» — в 5 классе 12 раз): верхние (зелёные) — "
              "учебники биологии, нижние (синие) — все учебники прошлых классов.")
    gl = wrap(draw, legend, font(14), width)
    for i, ln in enumerate(gl):
        draw.text((x0, CANVAS[1] - MARGIN - 12 - (len(gl) - i) * 18), ln, font=font(14), fill=GREY)
    return canvas


DENSITY_BLOCK_H = 215
ZONES = (("низкая", (86, 150, 214)), ("средняя", (110, 180, 90)),
         ("повышенная", (240, 160, 40)), ("высокая", (215, 45, 45)))


def zone_of(value: float, zones: tuple) -> int:
    q25, q75, q90 = zones
    return 0 if value < q25 else 1 if value < q75 else 2 if value < q90 else 3


def draw_density(draw: ImageDraw.ImageDraw, x: int, y: int, width: int, d: dict) -> None:
    """
    Терминологическая плотность страницы и где она на распределении по всем
    страницам учебника: гистограмма раскрашена по зонам (четверть самых
    разреженных страниц — низкая, середина — средняя, 75–90 % — повышенная,
    10 % самых плотных — высокая), маркер — эта страница.
    """
    vals, zones, value = d["all"], d["zones"], d["value"]
    draw.line((x, y - 6, x + width, y - 6), fill=(205, 205, 200), width=1)
    title_f, f, small = font(19, True), font(17), font(13)
    draw.text((x, y), "Терминологическая плотность", font=title_f, fill=INK)
    formula = "= терминов / слов × 100"
    draw.text((x + draw.textlength("Терминологическая плотность ", font=title_f), y + 3), formula,
              font=font(14), fill=GREY)
    y += 28
    if value is None:
        msg = ("Основного текста параграфа на странице нет (вне параграфа или после «Тренируемся») — "
               "плотность не считается.")
        for ln in wrap(draw, msg, f, width):
            draw.text((x, y), ln, font=f, fill=GREY)
            y += 23
    else:
        z = zone_of(value, zones)
        name, color = ZONES[z]
        share = sum(1 for v in vals if v < value) / len(vals) if vals else 0
        draw.rounded_rectangle((x, y, x + 132, y + 26), radius=6, fill=color)
        draw.text((x + 66 - draw.textlength(name, font=font(16, True)) / 2, y + 4), name,
                  font=font(16, True), fill="white")
        txt = (f"{value:.1f} терминов на 100 слов: {d['terms']} / {d['words']} × 100; "
               f"выше, чем на {share:.0%} страниц учебника")
        if d.get("approx"):
            txt += f". Основного текста мало (меньше {MIN_PAGE_WORDS} слов) — оценка приблизительная"
        lines = wrap(draw, txt, f, width - 145)
        for i, ln in enumerate(lines):
            draw.text((x + 145, y + 3 + i * 21), ln, font=f, fill=INK)
        y += max(30, len(lines) * 21 + 6)
    # Гистограмма распределения по страницам
    if not vals:
        return
    hx, hy, hw, hh = x + 4, y + 30, width - 8, 74
    top = max(vals)
    bins = 28
    step = top / bins if top else 1
    counts = [0] * bins
    for v in vals:
        counts[min(bins - 1, int(v / step))] += 1
    peak = max(counts) or 1
    bw = hw / bins
    for i, c in enumerate(counts):
        center = (i + 0.5) * step
        col = ZONES[zone_of(center, zones)][1]
        h = c / peak * (hh - 4)
        draw.rectangle((hx + i * bw + 1, hy + hh - h, hx + (i + 1) * bw - 1, hy + hh), fill=col)
    draw.line((hx, hy + hh, hx + hw, hy + hh), fill=(120, 120, 120), width=1)
    for v in (0, top / 2, top):
        lab = f"{v:.0f}"
        px = hx + (v / top if top else 0) * hw
        draw.text((px - draw.textlength(lab, font=small) / 2, hy + hh + 3), lab, font=small, fill=GREY)
    draw.text((hx + hw - draw.textlength("терминов на 100 слов", font=small), hy + hh + 17),
              "терминов на 100 слов", font=small, fill=GREY)
    draw.text((hx, hy + hh + 17), "число страниц учебника", font=small, fill=GREY)
    if value is not None:
        px = hx + min(1.0, value / top if top else 0) * hw
        draw.line((px, hy - 4, px, hy + hh), fill=INK, width=3)
        draw.polygon([(px - 7, hy - 12), (px + 7, hy - 12), (px, hy - 3)], fill=INK)
        lab = "эта страница"
        lx = min(max(hx, px - draw.textlength(lab, font=small) / 2), hx + hw - draw.textlength(lab, font=small))
        draw.text((lx, hy - 27), lab, font=small, fill=INK)


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
        f"  • определение дано на {LATE_MIN_PAGES}+ страниц позже первого упоминания в основном тексте (для новых "
        f"слов и терминов, которые ученик может не знать) — {kinds['late']}",
        f"  • есть только пояснение (синоним, «т. е.», происхождение слова) — {kinds['explain']}",
        f"  • термин без определения, который ученик может не помнить (не встречался, встречался редко "
        f"или {LONG_GAP_YEARS}+ лет назад) — {kinds['old']}",
        f"  • слово снова встречается после перерыва в {FAR_PAGES}+ страниц — {kinds['far']}",
        f"  • много новых слов на одной странице ({MANY_NEW_ON_PAGE}+) — {kinds['many']}",
        f"  • высокая плотность терминов на странице (10 % самых плотных) — {kinds['dense']}",
        f"  • параграф с высокой плотностью терминов (25 % самых плотных) — {kinds['sec']}",
        "",
        "Внизу справа на каждой странице — терминологическая плотность и гистограмма плотности по всем "
        "страницам учебника с отметкой этой страницы.",
        "Формула: плотность = Т / С × 100, где Т — число выделенных терминов на странице (из всех словарей; "
        "составной термин — одно вхождение, слово внутри более длинного термина отдельно не считается), "
        "С — число слов на странице. Учитывается только основной текст параграфа: от начала параграфа до "
        "рубрики «Тренируемся»; вопросы, «Это интересно», «Шаги к успеху», введение и приложения не входят. "
        "Пример (с. 78): 81 термин на 234 словах — 81 / 234 × 100 = 34,6 термина на 100 слов. Если на странице "
        f"меньше {MIN_PAGE_WORDS} слов основного текста, значение приблизительное и в распределение не входит.",
        "Цвет — зона по распределению страниц: синяя — низкая (четверть самых разреженных страниц), "
        "зелёная — средняя, оранжевая — повышенная (от 75 до 90 %), красная — высокая (10 % самых плотных).",
        "",
        f"Общеупотребительные термины (во всех учебниках прошлых классов встречались {EVERYDAY_TOTAL}+ раз, из "
        f"них {EVERYDAY_PRIMARY}+ раз ещё в начальной школе: «человек», «жизнь», «наблюдение» и т. п.) знакомы "
        f"почти всем, поэтому замечаний по ним нет; в плотности терминов они учитываются. "
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
            words = R["page_words"].get(p, 0)
            value = R["page_density"].get(p)
            approx = value is None and words > 0
            if approx:
                # Мало основного текста (подписи, рисунки): плотность считается, но помечается как приблизительная
                value = R["page_terms"].get(p, 0) / words * 100
            density = {"value": value, "all": sorted(R["page_density"].values()), "zones": R["density_zones"],
                       "terms": R["page_terms"].get(p, 0), "words": words, "approx": approx}
            img = render_page(src, pdf, p, marks.get(p, []), recs.get(p, []), header, notes,
                              a.get("term_sources"), density)
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
