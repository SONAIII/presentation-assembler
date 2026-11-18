#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""preflight.py — проверка окружения одной командой (онбординг, см. TESTING.md).

    python3 scripts/preflight.py            # все проверки, человеческий вывод
    python3 scripts/preflight.py --offline  # не ходить в сеть (пин не сверяется)
    python3 scripts/preflight.py --render   # + рендер-проба: соберёт крошечный
                                            # pptx и прогонит через PowerPoint
                                            # (запустит PP; ловит неактивированную
                                            # лицензию и запрет Automation ДО
                                            # первой настоящей сборки)

Проверяет всё, что TESTING.md перечисляет как пререквизиты: платформу,
python-пакеты, шрифты YS, PowerPoint, доступность шаблона (дев-файл → кэш →
Я.Диск) и кэш. Ничего не скачивает и не ставит — только диагностирует и
говорит, какой командой чинить.

Коды выхода: 0 — можно собирать (⚠ допустимы), 1 — есть блокеры (✗).
"""
from __future__ import annotations

import argparse
import importlib
import os
import platform
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.realpath(__file__))), "lib"))
import fetch_template as FT                     # noqa: E402
import template_profile as TP                   # noqa: E402  (пер-шаблонные шрифты/канва)

# Windows-консоль cp1251 (хэндофф Ultima 2026-08-11, O6.1): '✓' и '→' в неё
# не кодируются — preflight падал UnicodeEncodeError ДО проверки шрифтов,
# то есть на Windows не отвечал на главный вопрос. Двойная страховка:
# невыразимые символы → '?', а сами марки — ASCII-фолбэком.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(errors="replace")
        except (OSError, ValueError):
            pass


def _console_takes(s: str) -> bool:
    try:
        s.encode(getattr(sys.stdout, "encoding", None) or "utf-8")
        return True
    except (UnicodeEncodeError, LookupError):
        return False


OK, WARN, FAIL = ("✓", "⚠", "✗") if _console_takes("✓⚠✗") else ("[ok]", "[!]", "[X]")
EXPECTED_PPTX = "1.0.2"


class Report:
    def __init__(self):
        self.rows: list[tuple[str, str, str]] = []

    def add(self, mark: str, name: str, detail: str = ""):
        self.rows.append((mark, name, detail))
        print(f"  {mark} {name}" + (f" — {detail}" if detail else ""))

    @property
    def blockers(self) -> int:
        return sum(1 for m, _, _ in self.rows if m == FAIL)

    @property
    def warnings(self) -> int:
        return sum(1 for m, _, _ in self.rows if m == WARN)


def check_platform(rep: Report):
    sysname = platform.system()
    if sysname == "Darwin":
        rep.add(OK, "платформа", f"macOS {platform.mac_ver()[0]}")
    elif sysname == "Windows":
        rep.add(OK, "платформа", "Windows: сборка и гейты работают; "
                "пиксель-приёмка — PowerPoint через COM-экспорт "
                "(scripts/render.py), скрининг — LibreOffice")
    else:
        rep.add(OK, "платформа", f"{sysname}: сборка, гейты и LO-скрининг "
                "работают; пиксель-приёмка (PowerPoint) — macOS/Windows "
                "(quality-gates.md §4)")


def check_python(rep: Report):
    v = sys.version_info
    if (v.major, v.minor) >= (3, 11):
        rep.add(OK, "python", platform.python_version())
    else:
        rep.add(FAIL, "python", f"{platform.python_version()} — нужен 3.11+")


def check_packages(rep: Report):
    try:
        import pptx
        if pptx.__version__ == EXPECTED_PPTX:
            rep.add(OK, "python-pptx", pptx.__version__)
        else:
            rep.add(WARN, "python-pptx", f"{pptx.__version__} (движок обкатан "
                    f"на {EXPECTED_PPTX}; при странностях сборки — поставьте её: "
                    f"pip3 install python-pptx=={EXPECTED_PPTX})")
    except ImportError:
        rep.add(FAIL, "python-pptx", "не установлен — pip3 install "
                f"python-pptx=={EXPECTED_PPTX} (без него сборка не работает)")
    try:
        importlib.import_module("PIL")
        rep.add(OK, "Pillow", "есть (глиф-замеры lint)")
    except ImportError:
        rep.add(FAIL, "Pillow", "не установлен — pip3 install Pillow "
                "(без него не работает lint по глифам)")
    try:
        importlib.import_module("fitz")
        rep.add(OK, "PyMuPDF", "есть (PDF→PNG для рендера)")
    except ImportError:
        rep.add(WARN, "PyMuPDF", "не установлен — pip3 install PyMuPDF "
                "(сборка работает, но scripts/render.py не сможет отдать PNG)")
    check_custom_deps(rep)


def check_custom_deps(rep: Report):
    """Зависимости кастом-ветки (family:custom, references/custom-html.md) —
    без них обычная сборка работает, кастомные слайды не соберутся. Отдельная
    функция ради быстрого чека `--custom-only` (0.21.0): агент зовёт его на
    этапе ПЛАНА (SKILL.md шаг 2) и при отсутствии браузера планирует дек без
    custom, а не падает в конце сборки."""
    try:
        importlib.import_module("uharfbuzz")
        rep.add(OK, "uharfbuzz", "есть (PP-метрики custom-слайдов)")
    except ImportError:
        rep.add(WARN, "uharfbuzz", "не установлен — pip3 install uharfbuzz "
                "(нужен только для family:custom)")
    node = shutil.which("node")
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "lib"))
    import chrome as _chrome                     # единый резолвер (env + pw-browsers)
    chrome = _chrome.find_chrome()
    if node and chrome:
        rep.add(OK, "node + Chromium-браузер",
                f"есть (обмер HTML custom-слайдов): {chrome}")
    else:
        miss = " и ".join(x for x, ok in
                          (("node", node), ("Chromium-браузер", chrome)) if not ok)
        rep.add(WARN, "node + Chromium-браузер",
                f"нет: {miss} — family:custom не соберётся (остальной сборке "
                f"не мешает). Быстрый выход: npx playwright install chromium "
                f"(~150 МБ, найдётся автоматически) либо CHROME_PATH=<путь к "
                f"Chrome/Chromium>. Яндекс.Браузер НЕ годится: его headless "
                f"виснет (живая проба 2026-07-23)")


def _find_soffice() -> str | None:
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    for c in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice"):
        if os.path.isfile(c):
            return c
    return None


def check_powerpoint(rep: Report):
    so = _find_soffice()
    if platform.system() == "Windows":
        sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
        import render as R
        pp = R.find_powerpnt_win()
        if pp:
            rep.add(OK, "PowerPoint", "установлен — пиксель-приёмка через "
                    "COM-экспорт (нужен pywin32: pip install pywin32)")
        elif so:
            rep.add(WARN, "PowerPoint", "не найден — рендер уйдёт в LibreOffice "
                    "(скрининг: дефекты из lo-render-quirks.md перепроверять "
                    "мерой)")
        else:
            rep.add(WARN, "PowerPoint", "не найден, LibreOffice тоже нет — "
                    "превью не отрендерить, самопроверка только мерой")
        return
    if platform.system() != "Darwin":
        if so:
            rep.add(OK, "рендер", "LibreOffice найден — скрининг-рендер работает "
                    "(дефекты из lo-render-quirks.md перепроверять мерой; "
                    "пиксель-приёмка = PowerPoint на macOS/Windows)")
        else:
            rep.add(WARN, "рендер", "LibreOffice не найден — превью не "
                    "отрендерить, самопроверка только мерой; поставьте: "
                    "apt install libreoffice / brew install --cask libreoffice")
        return
    if os.path.isdir("/Applications/Microsoft PowerPoint.app"):
        rep.add(OK, "PowerPoint", "установлен (пиксель-приёмка; должен быть "
                "хоть раз запущен — лицензия)")
    elif so:
        rep.add(WARN, "PowerPoint", "не найден в /Applications — рендер уйдёт в "
                "LibreOffice (скрининг: дефекты из lo-render-quirks.md "
                "перепроверять мерой); для пиксель-приёмки поставьте PowerPoint")
    else:
        rep.add(WARN, "PowerPoint", "не найден, LibreOffice тоже нет — превью не "
                "отрендерить, самопроверка только мерой (qlmanage для приёмки "
                "запрещён, common-mistakes.md)")


def _zip_font_names(tdir: str) -> list[str]:
    """Имена шрифтов из fonts-ys.zip (локальный/кэш), если он уже на диске."""
    man = FT.load_manifest()
    fz = FT._entry(man, tdir)["fonts"]
    for p in (os.path.join(FT.SKILL_DIR, "Templates", tdir, fz["name"]),
              os.path.join(FT.cache_root(), tdir, fz["sha256"], fz["name"])):
        if os.path.exists(p):
            with zipfile.ZipFile(p) as z:
                return sorted({i.filename.replace("\\", "/").rsplit("/", 1)[-1]
                               for i in z.infolist()
                               if i.filename.lower().endswith((".ttf", ".otf"))})
    return []


def check_fonts(rep: Report, tdir: str):
    if platform.system() == "Windows":
        # автоустановки на Windows НЕТ (файл без записи в реестре не активируется);
        # ставится руками через проводник — ОС регистрирует шрифт сама
        if FT._win_fonts_present(tdir):
            rep.add(OK, "шрифты YS", "найдены (установлены руками)")
        else:
            link = FT.public_key(FT.load_manifest())
            rep.add(WARN, "шрифты YS", "НЕ установлены — PowerPoint молча подменит "
                    f"типографику. Руками: скачай fonts-ys.zip с {link} , распакуй, "
                    "выдели все шрифты → ПКМ → «Установить»")
        return
    dest = FT._font_dest()
    want = _zip_font_names(tdir)
    if want:
        missing = [f for f in want if not os.path.exists(os.path.join(dest, f))]
        if not missing:
            rep.add(OK, "шрифты YS", f"{len(want)} файлов в {dest}")
        else:
            rep.add(WARN, "шрифты YS", f"не хватает {len(missing)} из "
                    f"{len(want)} ({', '.join(missing[:3])}…) — поставятся сами "
                    "при первой сборке, руками: "
                    "python3 scripts/fetch_template.py --fonts")
        return
    # префиксы имён шрифтов — из профиля шаблона (fonts.files), не хардкод «ys text/geo»
    try:
        prefixes = tuple(sorted({v.split("-")[0].strip().lower()
                                 for v in TP.load(tdir)["fonts"]["files"].values()}))
    except (FileNotFoundError, KeyError):
        prefixes = ("ys text", "ys geo")        # шаблон без профиля кастома — прежний дефолт
    have = [f for f in os.listdir(dest)
            if f.lower().startswith(prefixes)] if os.path.isdir(dest) else []
    if have:
        rep.add(OK, "шрифты YS", f"найдены в {dest} ({len(have)} файлов)")
    else:
        rep.add(WARN, "шрифты YS", f"в {dest} не найдены — поставятся сами при "
                "первой сборке (нужна сеть), руками: "
                "python3 scripts/fetch_template.py --fonts")


def check_template(rep: Report, tdir: str, offline: bool):
    man = FT.load_manifest()
    tpl = FT._entry(man, tdir)["template"]
    local = os.path.join(FT.SKILL_DIR, "Templates", tdir, tpl["name"])
    cached = os.path.join(FT.cache_root(), tdir, tpl["sha256"], tpl["name"])
    if os.path.exists(local):
        rep.add(OK, "шаблон", f"дев-файл {os.path.relpath(local, FT.SKILL_DIR)}")
        src = "local"
    elif os.path.exists(cached):
        rep.add(OK, "шаблон", f"в кэше по SHA-256 пину {tpl['sha256'][:8]}…")
        src = "cache"
    else:
        src = "remote"
    if offline:
        if src == "remote":
            rep.add(FAIL, "шаблон", "локально нет, а --offline: скачайте "
                    f"{tpl['name']} по ссылке из Templates/manifest.json и "
                    f"положите в Templates/{tdir}/")
        else:
            rep.add(WARN, "пин Я.Диска", "офлайн — сверка пропущена")
        return
    try:
        entry = FT._entry(man, tdir)
        # пер-шаблонный public_key + канонная склейка пути (пустой remote_dir =
        # корень полки) — как в самом fetch; глобальный ключ + ручная склейка
        # давали ложный HTTPError для gor-services (полка своя, remote_dir="")
        pk = FT.public_key(man, entry)
        meta = FT._remote_meta(pk, FT._remote_path(entry, tpl["name"]))
        pin_ok = meta.get("md5") == tpl["md5"] and meta.get("size") == tpl["size"]
        if src == "remote":
            if pin_ok:
                rep.add(OK, "шаблон", f"локально нет — скачается при первой "
                        f"сборке (~{tpl['size'] >> 20}МБ, пин на Диске сходится)")
            else:
                rep.add(FAIL, "шаблон", "локально нет, а ревизия на Диске "
                        "разошлась с пином манифеста — сообщите владельцу скилла")
        elif pin_ok:
            rep.add(OK, "пин Я.Диска", "ревизия на полке совпадает с манифестом")
        else:
            rep.add(WARN, "пин Я.Диска", "полку обновили без бампа пина — "
                    "сборке не мешает (движок берёт локальный/кэш), но владельцу "
                    "скилла пора перекалибровать — сообщите команде (тикет в VISCOMMSSKILLS)")
    except Exception as e:  # noqa: BLE001 — сеть недоступна/флап
        mark = FAIL if src == "remote" else WARN
        rep.add(mark, "сеть до Я.Диска", f"не достучался ({e.__class__.__name__})"
                + ("" if src != "remote" else " — а шаблона локально нет; "
                   "скачайте руками по ссылке из Templates/manifest.json"))


def check_render_probe(rep: Report):
    """Крошечный pptx → PowerPoint → PNG. Единственная проверка, которая
    ЗАПУСКАЕТ PowerPoint, поэтому только под флагом --render. Ловит:
    неактивированную лицензию, невыданное разрешение Automation, PP-зомби."""
    sysname = platform.system()
    if sysname not in ("Darwin", "Windows"):
        rep.add(WARN, "рендер-проба", "пропущена: PowerPoint-рендер есть только "
                "на macOS и Windows")
        return
    if sysname == "Darwin" \
            and not os.path.isdir("/Applications/Microsoft PowerPoint.app"):
        rep.add(WARN, "рендер-проба", "пропущена: PowerPoint не установлен")
        return
    try:
        from pptx import Presentation
        from pptx.util import Inches
    except ImportError:
        rep.add(WARN, "рендер-проба", "пропущена: нет python-pptx")
        return
    import shutil
    import render as R
    dpi = 110
    work = tempfile.mkdtemp(prefix="preflight_render_",
                            dir="/private/tmp" if os.path.isdir("/private/tmp")
                            else None)             # TCC-грабля — только macOS
    probe = os.path.join(work, "probe.pptx")
    try:
        prs = Presentation()                       # дефолтная канва 10×7.5"
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        box.text_frame.text = "preflight render probe"
        prs.save(probe)
        if sysname == "Windows":
            if not R.find_powerpnt_win():
                rep.add(WARN, "рендер-проба", "пропущена: PowerPoint не найден "
                        "в штатных путях Office")
                return
            pngs = R.via_powerpoint_com(probe, work, dpi)
        else:
            pngs = R.via_powerpoint(probe, work, dpi)
        if not pngs:
            rep.add(WARN, "рендер-проба", "PowerPoint не отдал PDF — запустите "
                    "его один раз руками (активация лицензии / диалоги первого "
                    "запуска) и повторите --render")
            return
        from PIL import Image
        w, h = Image.open(pngs[0]).size
        want = round(10 * dpi)
        if abs(w - want) <= 2:
            rep.add(OK, "рендер-проба", f"PowerPoint отрендерил тестовый слайд "
                    f"({w}×{h}px @ {dpi}dpi)")
        else:
            rep.add(WARN, "рендер-проба", f"PNG шириной {w}px вместо ~{want}px — "
                    "похоже, рендерил не PowerPoint (см. каскад в "
                    "quality-gates.md)")
    except Exception as e:  # noqa: BLE001 — важна причина, не тип
        rep.add(WARN, "рендер-проба", f"не прошла ({e.__class__.__name__}) — "
                "частые причины: не выдано разрешение Automation "
                "(Настройки → Конфиденциальность → Автоматизация → разрешить "
                "терминалу управлять PowerPoint), PP ни разу не запускался, "
                "PP-зомби (перезапустите его)")
    finally:
        shutil.rmtree(work, ignore_errors=True)


def check_self_contained(rep: Report):
    """Самодостаточность папки: файлы, которые ДОЛЖНЫ быть внутри скилла
    (иначе кастом-ветка полезет наружу — чего быть не должно). Пер-шаблонно:
    у каждого шаблона манифеста с кастом-контуром (custom-profile.json) обязан
    быть свой assets/<tdir>/deck.css."""
    skill = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    checked = False
    for tdir in FT.load_manifest().get("templates", {}):
        if not os.path.exists(os.path.join(skill, "layouts", tdir, "custom-profile.json")):
            continue                            # у шаблона нет кастом-контура — CSS не нужен
        checked = True
        css = os.path.join(skill, "assets", tdir, "deck.css")
        if os.path.exists(css):
            rep.add(OK, "дизайн-система кастома", f"assets/{tdir}/deck.css в скилле")
        else:
            rep.add(WARN, "дизайн-система кастома", f"нет assets/{tdir}/deck.css — "
                    f"family:custom ({tdir}) не соберётся; папка неполная")
    if not checked:
        rep.add(OK, "дизайн-система кастома", "шаблонов с кастом-контуром нет")


def check_cache(rep: Report):
    root = FT.cache_root()
    try:
        os.makedirs(root, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=root):
            pass
        rep.add(OK, "кэш", root)
    except OSError as e:
        rep.add(FAIL, "кэш", f"{root} не пишется ({e}) — переопределите "
                "PRESENTATION_ASSEMBLER_CACHE")


def main() -> int:
    ap = argparse.ArgumentParser(description="проверка окружения скилла")
    ap.add_argument("--template", default="eda", help="tdir шаблона (еда|eda)")
    ap.add_argument("--offline", action="store_true", help="не ходить в сеть")
    ap.add_argument("--render", action="store_true",
                    help="рендер-проба: тестовый pptx через PowerPoint "
                         "(запустит PP на несколько секунд)")
    ap.add_argument("--custom-only", action="store_true",
                    help="быстрый чек только кастом-контура (node + Chromium-"
                         "браузер + uharfbuzz, <1 c) — для этапа плана "
                         "(SKILL.md шаг 2); exit 0 = custom доступен, 2 = нет")
    a = ap.parse_args()
    tdir = FT.TDIR_ALIASES.get(a.template.lower(), a.template.lower())

    if a.custom_only:
        print("preflight (кастом-контур):")
        rep = Report()
        check_custom_deps(rep)
        print()
        ok = not rep.warnings and not rep.blockers
        print(f"{OK} custom доступен" if ok else
              f"{FAIL} custom НЕДОСТУПЕН — планируй дек из готовых семейств "
              f"или дай браузер (строки выше)")
        return 0 if ok else 2

    print("preflight presentation-assembler:")
    rep = Report()
    check_platform(rep)
    check_python(rep)
    check_packages(rep)
    check_powerpoint(rep)
    check_fonts(rep, tdir)
    check_template(rep, tdir, a.offline)
    check_self_contained(rep)
    check_cache(rep)
    if a.render:
        check_render_probe(rep)

    print()
    if rep.blockers:
        print(f"{FAIL} {rep.blockers} блокер(ов), {rep.warnings} замечаний — "
              "сборка не взлетит, чините по строкам выше")
        return 1
    if rep.warnings:
        print(f"{OK} можно собирать; {rep.warnings} замечание(й) выше — "
              "сборке не мешают")
    else:
        print(f"{OK} окружение полностью готово")
    return 0


if __name__ == "__main__":
    sys.exit(main())
