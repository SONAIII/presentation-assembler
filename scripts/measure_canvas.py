#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""measure_canvas.py — обмер дизайнерской сетки шаблона → layouts/<tdir>/_canvas.json.

Зачем: движок раскладки lib/layout.py кладёт GENERATE-композиции в КАНВУ —
измеренную сетку шаблона (заголовочный бокс + content_band + поля + гаттер).
Числа нельзя брать «на глаз»: дизайнер двигает полосу контента и футер между
ревизиями, а спеки/атомы калиброваны под конкретную геометрию. Скрипт снимает
сетку машинно — по МОДЕ (самому частому округлённому значению) измерений по
НЕскрытым слайдам — и пишет контракт, который читает lib/layout.py (load_adapter).

Метод (одно измерение — одна мода; гистограммы в _evidence для ручной сверки):
  * heading      — бокс плейсхолдера-заголовка (idx 0, фолбэк по имени «Title 1»):
                   top/left/w/h, мода по слайдам;
  * content_band — конверт реального контента: left = мода min-left по слайду,
                   right = мода max-right, bottom = мода max-bottom (фон,
                   полноэкранные фото и декоративные полосы отсеяны по площади/
                   размаху); top — мода ВЕРХОВ «больших» блоков (чарт/панель/карта)
                   ниже заголовка, уточнённая средним выигравшего кластера;
  * margins      — left = content_band.left; right = slide_w − content_band.right;
  * gutter       — мода МИНИМАЛЬНОГО зазора между колонками (широкие шейпы с
                   вертикальным перекрытием) по многоколоночным слайдам. Самый
                   зыбкий замер (в шаблоне Еды зазоры шумные) — берём тесный
                   повторяющийся гаттер; фолбэк 0.78 (его же дефолтит layout.py).

Скрытые слайды (p:sld show="0") пропускаются — это палитра атомов/черновики
(тот же тест, что в scripts/palette_lint.py). Обмер по ВЕРХНЕУРОВНЕВЫМ шейпам:
в группы не заходим — у группы свой bbox, вложенные члены дублировали бы конверт.

  python3 scripts/measure_canvas.py                  # обмерять Еду, писать _canvas.json
  python3 scripts/measure_canvas.py --template eda   # то же явно (алиас «еда» ок)
  python3 scripts/measure_canvas.py --template путь/шаблон.pptx   # произвольный .pptx
  python3 scripts/measure_canvas.py --check          # сверить с текущим _canvas.json, НЕ писать
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)          # чтобы import fetch_template (как в inventory_template)

EMU = 914400                      # EMU в дюйме (тот же множитель, что в lib/*, inventory)

# _doc держим байт-в-байт с существующим _canvas.json — снимок коммитится, дифф
# должен дребезжать только на изменившихся ЧИСЛАХ, не на служебной строке.
CANVAS_DOC = (
    "Измеренная сетка шаблона (scripts/measure_canvas.py). Корневой контейнер "
    "движка lib/layout.py: GENERATE-композиции живут в content_band, стиль — из "
    "atoms.json, шрифты/цвета — brand-tokens.json."
)

# --- пороги отсева (все — доли слайда / дюймы; вынесены, чтобы были видимы) ---
BG_AREA_FRAC = 0.60      # площадь ≥ 60% слайда → фон / полноэкранное фото, не контент
BG_SPAN_FRAC = 0.92      # почти во всю ширину ИЛИ высоту → декоративная полоса / бок-фото
BIG_AREA_FRAC = 0.12     # «большой» блок (чарт/панель/карта): его верх = старт контента
COL_MIN_W_IN = 3.0       # колонкой для гаттера считаем только достаточно широкий шейп
COL_VOVERLAP_IN = 1.0    # ... с вертикальным перекрытием не меньше этого
COL_GAP_MAX_IN = 3.0     # зазор шире — это не гаттер, а разрыв макета
GUTTER_FALLBACK_IN = 0.78
CHECK_TOL_IN = 0.05      # допуск сверки --check по дюймовым полям


# --------------------------------------------------------------- предикаты шейпа

def _visible(slide) -> bool:
    """Слайд показывается (не hidden). Скрытый p:sld несёт show="0"
    (тот же тест, что scripts/palette_lint.py:collect_palette)."""
    return slide._element.get("show") != "0"


def _is_title(sh) -> bool:
    return bool(sh.is_placeholder and sh.placeholder_format.idx == 0)


def _is_footer(sh) -> bool:
    return bool(sh.is_placeholder and sh.placeholder_format.idx == 10)


def _layout_ftr_top(slide):
    """Верх штатного футера PowerPoint (type=ftr), унаследованного с лэйаута.
    Нужен шаблонам, где контентные слайды НЕ несут свой футер-шейп, а нижняя
    полоса заведена футером лэйаута (gor-services: ftr idx=3 @14.14; контентные
    слайды наследуют его и обмер по slide.shapes его не видит). Еда футеров
    type=ftr не имеет → для неё фолбэк — no-op, обмер байт-в-байт прежний.
    Возвращает top в дюймах или None."""
    try:
        for sh in slide.slide_layout.placeholders:
            if sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                g = _geom_in(sh)
                if g is not None:
                    return g[1]
    except Exception:  # noqa: BLE001 — редкие лэйауты без сериализуемых плейсхолдеров
        pass
    return None


def _is_slidenum(sh) -> bool:
    """Плейсхолдер номера страницы (build_deck.py тем же тестом ловит поле номера)."""
    try:
        return 'type="slidenum"' in sh._element.xml
    except Exception:  # noqa: BLE001 — редкие шейпы без сериализуемого xml
        return False


def _title_shape(slide):
    """Плейсхолдер-заголовок: idx 0; фолбэк — верхнеуровневый шейп «Title 1»
    (build_deck.fill_heading заливает заголовок именно по этому имени)."""
    for sh in slide.shapes:
        if _is_title(sh):
            return sh
    for sh in slide.shapes:
        if sh.name == "Title 1":
            return sh
    return None


def _geom_in(sh):
    """(left, top, w, h) шейпа в дюймах или None, если геометрия не задана
    (наследуемый плейсхолдер без своего xfrm)."""
    if None in (sh.left, sh.top, sh.width, sh.height):
        return None
    return (sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU)


def _is_background(l, t, w, h, sw, sh_) -> bool:
    """Фон / полноэкранное фото / декоративная полоса — из конверта контента вон.
    Именно этот отсев поднимает моды до 24.96 / 14.09 (иначе выигрывают
    полноэкранные 26.66 / 15.0)."""
    if (w * h) / (sw * sh_) >= BG_AREA_FRAC:
        return True
    if w >= BG_SPAN_FRAC * sw:
        return True
    if h >= BG_SPAN_FRAC * sh_:
        return True
    return False


def _content_shapes(slide, sw, sh_):
    """Верхнеуровневые контент-шейпы слайда: без заголовка/футера/номера, без
    фонов, с ненулевой геометрией. Возвращает [(l, t, w, h, area_frac)]."""
    out = []
    for sh in slide.shapes:                      # только верхний уровень, в группы не идём
        if _is_title(sh) or _is_footer(sh) or _is_slidenum(sh):
            continue
        g = _geom_in(sh)
        if g is None:
            continue
        l, t, w, h = g
        if w <= 0 or h <= 0 or _is_background(l, t, w, h, sw, sh_):
            continue
        out.append((l, t, w, h, (w * h) / (sw * sh_)))
    return out


# ------------------------------------------------------------------ утилиты моды

def _mode(counter: Counter, prefer: str = "max"):
    """Самое частое значение (None на пустом счётчике). При равном счёте —
    детерминированный тай-брейк по принципу «канва = конверт контента»
    (низ/право → большее, верх/лево → prefer='min'), а НЕ порядок вставки:
    иначе перестановка слайдов в шаблоне меняла бы контракт канвы."""
    if not counter:
        return None
    if prefer == "min":
        return max(counter.items(), key=lambda kv: (kv[1], -kv[0]))[0]
    return max(counter.items(), key=lambda kv: (kv[1], kv[0]))[0]


def _mode_or(counter: Counter, fallback, prefer: str = "max"):
    """Мода или фолбэк. Именно `is None`, не `or`: честная мода 0.0
    (контент впритык к краю) — легитимное значение, глотать её нельзя."""
    m = _mode(counter, prefer)
    return fallback if m is None else m


def _portable_path(path: str) -> str:
    """Путь в снимке — переносимый (относительный от корня скилла), как в
    inventory_template._portable_path: снимок коммитится как бейзлайн."""
    ap = os.path.abspath(path)
    return (os.path.relpath(ap, ROOT) if ap.startswith(ROOT + os.sep)
            else os.path.basename(ap))


# --------------------------------------------------------------------- обмер

def measure_canvas(template_path: str) -> dict:
    """Обмерять шаблон → dict в схеме _canvas.json (порядок ключей — как в файле)."""
    prs = Presentation(template_path)
    sw = prs.slide_width / EMU
    sh_ = prs.slide_height / EMU

    # 1) заголовок: мода бокса плейсхолдера-заголовка по видимым слайдам
    t_top, t_left, t_w, t_h = Counter(), Counter(), Counter(), Counter()
    # 2) конверт контента: по слайду — min-left, max-right, max-bottom
    c_left, c_right, c_bottom = Counter(), Counter(), Counter()
    # 2a) верх полосы футера-якоря: низ контентной полосы = ВЕРХ футера (footer
    #     неприкосновенен, контент до него, не в него). Сильная стабильная мода
    #     по всем слайдам с футером — в отличие от шумного max-bottom контента.
    f_top = Counter()
    # 3) верхи «больших» блоков (сырьё; порог по заголовку применим ниже)
    big_raw = []                 # [(top_raw)] среди контент-шейпов с area ≥ BIG_AREA_FRAC
    # 4) гаттер: по слайду — минимальный зазор между колонками
    gutter = Counter()

    n_measured, n_hidden = 0, 0
    for slide in prs.slides:
        if not _visible(slide):
            n_hidden += 1
            continue
        n_measured += 1

        title = _title_shape(slide)
        if title is not None:
            g = _geom_in(title)
            if g is not None:
                l, t, w, h = g
                t_top[round(t, 2)] += 1
                t_left[round(l, 2)] += 1
                t_w[round(w, 2)] += 1
                t_h[round(h, 2)] += 1

        found_footer = False
        for sh in slide.shapes:
            if _is_footer(sh):
                g = _geom_in(sh)
                if g is not None:
                    f_top[round(g[1], 2)] += 1
                found_footer = True
                break
        if not found_footer:                     # футер унаследован с лэйаута (type=ftr)
            ft = _layout_ftr_top(slide)
            if ft is not None:
                f_top[round(ft, 2)] += 1

        cs = _content_shapes(slide, sw, sh_)
        if cs:
            c_left[round(min(l for l, _, _, _, _ in cs), 2)] += 1
            c_right[round(max(l + w for l, _, w, _, _ in cs), 2)] += 1
            c_bottom[round(max(t + h for _, t, _, h, _ in cs), 2)] += 1
        for l, t, w, h, af in cs:
            if af >= BIG_AREA_FRAC:
                big_raw.append(t)

        # гаттер: тесный зазор между широкими вертикально-перекрытыми колонками
        cols = sorted([(l, t, w, h) for l, t, w, h, _ in cs if w >= COL_MIN_W_IN],
                      key=lambda r: r[0])
        slide_gaps = []
        for i, (l1, t1, w1, h1) in enumerate(cols):
            r1 = l1 + w1
            for j, (l2, t2, w2, h2) in enumerate(cols):
                if j == i or l2 < r1 - 0.05:
                    continue
                if min(t1 + h1, t2 + h2) - max(t1, t2) < COL_VOVERLAP_IN:
                    continue
                gap = l2 - r1
                if 0.1 < gap < COL_GAP_MAX_IN:
                    slide_gaps.append(gap)
        if slide_gaps:
            gutter[round(min(slide_gaps), 2)] += 1

    # --- сведение мод -------------------------------------------------------
    h_top = _mode_or(t_top, 0.0)
    h_left = _mode_or(t_left, 0.0)
    h_w = _mode_or(t_w, 0.0)
    h_h = _mode_or(t_h, 0.0)
    head_bottom = h_top + h_h

    # top контента: мода верхов «больших» блоков НИЖЕ заголовка; уточняем средним
    # выигравшего кластера (даёт 3 знака — как 3.691 в эталоне). Гистограмму
    # считаем на округлённых значениях, финал — среднее сырых в моде-корзине.
    big_below = [t for t in big_raw if t > head_bottom - 0.05]
    big_hist = Counter(round(t, 2) for t in big_below)
    if big_hist:
        win = _mode(big_hist, "min")
        bucket = [t for t in big_below if round(t, 2) == win]
        cb_top = round(sum(bucket) / len(bucket), 3)
    else:
        cb_top = round(head_bottom + 0.6, 3)      # деградация: нет больших блоков

    cb_left = _mode_or(c_left, round(h_left, 2), "min")
    cb_right = _mode_or(c_right, round(sw, 2))
    # низ полосы = верх футера-якоря (сильная мода); нет футера ни на одном
    # слайде → деградация на max-bottom контента с предупреждением
    footer_top = _mode(f_top, "min")
    if footer_top is not None:
        cb_bottom = footer_top
    else:
        cb_bottom = _mode_or(c_bottom, round(sh_, 2))
        top2 = c_bottom.most_common(2) if c_bottom else []
        weak = not top2 or top2[0][1] < 3 or (len(top2) > 1 and top2[1][1] == top2[0][1])
        if weak:
            print("⚠ футер-якорь не найден, низ канвы взят по слабой моде "
                  "контента — сверь глазами по шаблону, прежде чем коммитить")
    gut = _mode_or(gutter, GUTTER_FALLBACK_IN)

    canvas = {
        "_doc": CANVAS_DOC,
        "template_file": _portable_path(template_path),
        "slide_w_in": round(sw, 3),
        "slide_h_in": round(sh_, 3),
        "margin_left_in": round(cb_left, 3),
        "margin_right_in": round(sw - cb_right, 3),
        "heading": {
            "top_in": round(h_top, 2),
            "left_in": round(h_left, 2),
            "w_in": round(h_w, 2),
            "h_in": round(h_h, 2),
        },
        "content_band": {
            "left_in": round(cb_left, 2),
            "right_in": round(cb_right, 2),
            "top_in": cb_top,
            "bottom_in": round(cb_bottom, 2),
            "w_in": round(cb_right - cb_left, 3),
            "h_in": round(cb_bottom - cb_top, 3),
        },
        "gutter_in": round(gut, 2),
        "_evidence": {
            "slides_measured": n_measured,
            "hidden_skipped": n_hidden,
            "title_top_mode_hits": [list(x) for x in t_top.most_common(3)],
            "footer_top_mode_hits": [list(x) for x in f_top.most_common(3)],
            "content_bottom_mode_hits": [list(x) for x in c_bottom.most_common(3)],
            "content_right_mode_hits": [list(x) for x in c_right.most_common(3)],
            "big_shape_top_mode_hits": [list(x) for x in big_hist.most_common(5)],
        },
    }
    return canvas


# --------------------------------------------------------------------- резолв

def _resolve(arg: str):
    """(абсолютный путь к .pptx, tdir). arg — либо tdir (eda|еда), либо путь к .pptx."""
    import fetch_template  # локальный импорт: тянет сеть только если шаблона нет
    if arg.lower().endswith(".pptx"):
        path = os.path.abspath(arg)
        if not os.path.exists(path):
            raise SystemExit(f"measure_canvas: нет файла шаблона {path}")
        tdir = os.path.basename(os.path.dirname(path)) or "eda"
        return path, fetch_template.TDIR_ALIASES.get(tdir, tdir)
    tdir = fetch_template.TDIR_ALIASES.get(arg.lower(), arg.lower())
    return fetch_template.resolve_template(tdir), tdir


# --------------------------------------------------------------------- сверка

# дюймовые/целые поля контракта, которые сверяет --check (гистограммы _evidence
# по природе дребезжат — их не гейтим, только показываем).
_CHECK_FIELDS = [
    ("slide_w_in", "float"), ("slide_h_in", "float"),
    ("margin_left_in", "float"), ("margin_right_in", "float"),
    ("heading.top_in", "float"), ("heading.left_in", "float"),
    ("heading.w_in", "float"), ("heading.h_in", "float"),
    ("content_band.left_in", "float"), ("content_band.right_in", "float"),
    ("content_band.top_in", "float"), ("content_band.bottom_in", "float"),
    ("content_band.w_in", "float"), ("content_band.h_in", "float"),
    ("gutter_in", "float"),
    ("_evidence.slides_measured", "int"), ("_evidence.hidden_skipped", "int"),
]


def _dig(d: dict, dotted: str):
    cur = d
    for k in dotted.split("."):
        if not isinstance(cur, dict) or k not in cur:
            return None
        cur = cur[k]
    return cur


def do_check(fresh: dict, out_path: str) -> int:
    if not os.path.exists(out_path):
        print(f"[check] эталон {out_path} не найден — сверять не с чем "
              f"(запусти без --check, чтобы создать).")
        return 1
    with open(out_path, encoding="utf-8") as f:
        old = json.load(f)
    print(f"[check] сверка с {_portable_path(out_path)} "
          f"(допуск {CHECK_TOL_IN}\" по дюймовым полям)\n")
    drift = 0
    for field, kind in _CHECK_FIELDS:
        o, n = _dig(old, field), _dig(fresh, field)
        if kind == "int":
            ok = (o == n)
            delta = "" if ok else f"  Δ {o}→{n}"
        else:
            if o is None or n is None:
                ok, delta = False, f"  Δ {o}→{n}"
            else:
                ok = abs(float(o) - float(n)) <= CHECK_TOL_IN
                delta = "" if ok else f"  Δ {o}→{n} ({float(n) - float(o):+.3f})"
        drift += 0 if ok else 1
        print(f"  {'OK  ' if ok else 'DRIFT'}  {field:<28} {n}{delta}")
    print()
    if drift:
        print(f"[check] РАСХОЖДЕНИЕ в {drift} поле(ях) — шаблон сдвинулся против "
              f"_canvas.json; перезапусти measure_canvas без --check и перекалибруй "
              f"layouts/ (см. references/quality-gates.md).")
    else:
        print("[check] сетка совпадает с эталоном.")
    return 1 if drift else 0


# --------------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(
        description="обмер сетки шаблона (heading/content_band/поля/гаттер) → _canvas.json")
    ap.add_argument("--template", default="eda",
                    help="tdir шаблона (eda|еда) или путь к .pptx (по умолчанию eda)")
    ap.add_argument("--out", default=None,
                    help="куда писать (по умолчанию layouts/<tdir>/_canvas.json)")
    ap.add_argument("--check", action="store_true",
                    help="сверить свежий обмер с существующим _canvas.json, НЕ записывать")
    a = ap.parse_args()

    template_path, tdir = _resolve(a.template)
    out_path = a.out or os.path.join(ROOT, "layouts", tdir, "_canvas.json")

    canvas = measure_canvas(template_path)

    if a.check:
        return do_check(canvas, out_path)

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(canvas, f, ensure_ascii=False, indent=2)
    ev = canvas["_evidence"]
    band = canvas["content_band"]
    print(f"[measure] {_portable_path(template_path)}: "
          f"{ev['slides_measured']} слайдов обмерено, {ev['hidden_skipped']} скрытых пропущено")
    print(f"[measure] heading top={canvas['heading']['top_in']}\"  "
          f"content_band {band['w_in']}×{band['h_in']}\" "
          f"(top={band['top_in']} bottom={band['bottom_in']})  "
          f"gutter={canvas['gutter_in']}\"")
    print(f"[measure] JSON → {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
