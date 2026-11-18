#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Рендер .pptx в PNG по слайдам для визуальной самопроверки (не финальная выдача).

Usage:
    python3 scripts/render.py <pptx> [--dpi N] [-o/--out DIR] [--slides N,M]
      <pptx>      путь к .pptx (приводится к abspath; нет файла → return 1)
      --dpi N     разрешение PDF→PNG, default 150 (приёмка — 110, см. quality-gates.md §4)
      -o/--out    каталог PNG, дефолт <стем>.preview рядом с pptx
      --slides    растеризовать ТОЛЬКО эти слайды (1-based; PDF движки PP/LO
                  отдают целиком — экономится растеризация PDF→PNG)
    Пути готовых PNG печатаются в stdout по одному в строке (отсортированы).
    Запускать из /private/tmp (TCC/FDA-грабли macOS; правило — quality-gates.md §4,
    в самом render.py путь не ограничен). По одному деку за прогон, в фоне.

Каскад движков (первый давший PNG, затем break):
  1. PowerPoint (macOS + установлен) — AppleScript «save as PDF», «как в PowerPoint»,
     пиксель-точно на этой машине — эталонная приёмка;
  1а. PowerPoint (Windows + установлен) — COM-экспорт Open(ReadOnly) →
     SaveAs(pdf), та же пиксель-точность (хэндофф Ultima 2026-08-11, O6.4:
     раньше рендер POWERPNT.EXE не видел и честной приёмки на Windows не было);
  2. LibreOffice (soffice --headless --convert-to pdf) — скрининг на любой ОС;
  3. qlmanage (-s 1600) — ТОЛЬКО обложка, аварийный фолбэк; для приёмки НЕПРИГОДЕН.
PDF→PNG: PyMuPDF (fitz), фолбэк pdftoppm (poppler-utils); нет ни того ни другого → SystemExit.
"""
import argparse
import os
import platform
import shutil
import subprocess
import sys
import time


def pdf_to_png(pdf, outdir, dpi, only=None):
    """only — set 1-based номеров слайдов: растеризуются ТОЛЬКО они (фидбек
    2026-08-10 «--slides всё равно экспортирует весь дек»). Сам PDF движки
    (PowerPoint/LibreOffice) отдают целиком — выборочного экспорта у них нет,
    экономится растеризация (она и растёт с dpi и числом слайдов)."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return _pdf_to_png_poppler(pdf, outdir, dpi, only)
    out = []
    zoom = dpi / 72.0
    m = fitz.Matrix(zoom, zoom)
    with fitz.open(pdf) as doc:
        w = len(str(doc.page_count))
        for i, page in enumerate(doc, 1):
            if only and i not in only:
                continue
            png = os.path.join(outdir, f"slide-{str(i).zfill(w)}.png")
            page.get_pixmap(matrix=m).save(png)
            out.append(png)
        if only and not out:
            sys.stderr.write(f"render: --slides {sorted(only)}: таких страниц "
                             f"нет (в PDF {doc.page_count})\n")
    return out


def _pdf_to_png_poppler(pdf, outdir, dpi, only=None):
    """Фолбэк PDF→PNG без PyMuPDF (фидбек I, Linux/облако 2026-07-16):
    pdftoppm из poppler-utils решает то же самое; only — постранично -f/-l."""
    if not shutil.which("pdftoppm"):
        raise SystemExit("нет ни PyMuPDF, ни pdftoppm (poppler-utils) — "
                         "поставь одно из двух: pip install PyMuPDF | "
                         "apt install poppler-utils")
    prefix = os.path.join(outdir, "slide")
    if only:
        for i in sorted(only):
            subprocess.run(["pdftoppm", "-png", "-r", str(int(dpi)),
                            "-f", str(i), "-l", str(i), pdf, prefix],
                           check=True, timeout=600)
    else:
        subprocess.run(["pdftoppm", "-png", "-r", str(int(dpi)), pdf, prefix],
                       check=True, timeout=600)
    out = sorted(
        os.path.join(outdir, f) for f in os.listdir(outdir)
        if f.startswith("slide-") and f.endswith(".png"))
    return out


def via_powerpoint(pptx, outdir, dpi, only=None):
    pptx = os.path.abspath(pptx)
    if pptx.startswith("/private/tmp/"):
        pptx = "/tmp/" + pptx[len("/private/tmp/"):]
    pdf = os.path.splitext(pptx)[0] + ".pdf"  # песочница PowerPoint — PDF рядом с pptx
    subprocess.run(["open", "-n", "-a", "Microsoft PowerPoint", pptx],
                   check=True, timeout=30, stdout=subprocess.DEVNULL,
                   stderr=subprocess.DEVNULL)
    script = '''on run argv
 set pptxPath to item 1 of argv
 set pdfPath to item 2 of argv
 with timeout of 170 seconds
  tell application "Microsoft PowerPoint"
   set matches to every presentation whose full name is pptxPath
   if (count of matches) is 0 then error "test presentation not ready"
   set d to item 1 of matches
   save d in (POSIX file pdfPath) as save as PDF
   close d saving no
  end tell
 end timeout
end run'''
    last = None
    for _ in range(10):
        try:
            subprocess.run(["osascript", "-e", script, pptx, pdf], check=True, timeout=180,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            last = None
            break
        except subprocess.CalledProcessError as error:
            last = error
            time.sleep(1)
    if last is not None:
        raise last
    lock = os.path.join(os.path.dirname(pptx), "~$" + os.path.basename(pptx))
    if os.path.isfile(lock):
        try:
            os.remove(lock)
        except OSError:
            pass
    if not os.path.isfile(pdf):
        return []
    pngs = pdf_to_png(pdf, outdir, dpi, only)
    try:
        os.remove(pdf)
    except OSError:
        pass
    return pngs


def find_powerpnt_win():
    """POWERPNT.EXE в штатных путях Office (Click-to-Run root/OfficeNN и
    MSI OfficeNN, обе разрядности). COM-у путь не нужен (класс резолвится
    реестром) — проверка нужна каскаду, чтобы не дёргать COM без PowerPoint."""
    import glob as _glob
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env)
        if not base:
            continue
        for pat in (os.path.join(base, "Microsoft Office", "root", "Office*", "POWERPNT.EXE"),
                    os.path.join(base, "Microsoft Office", "Office*", "POWERPNT.EXE")):
            hits = _glob.glob(pat)
            if hits:
                return hits[0]
    return None


def via_powerpoint_com(pptx, outdir, dpi, only=None):
    """PowerPoint на Windows через COM (хэндофф Ultima 2026-08-11, O6.4 —
    рецепт приёмки тестера, работал стабильно и пиксель-точно):
    Presentations.Open(ReadOnly, WithWindow=False) → SaveAs(pdf, 32=ppSaveAsPDF)
    → общий pdf_to_png. ReadOnly бережёт файл, WithWindow не мигает окном и не
    трогает открытые документы пользователя. Нужен pywin32 или comtypes."""
    try:
        from win32com.client import Dispatch as _com
    except ImportError:
        try:
            from comtypes.client import CreateObject as _com
        except ImportError:
            sys.stderr.write("render: [PowerPoint COM] нужен pywin32 или "
                             "comtypes: pip install pywin32\n")
            return []
    pptx = os.path.abspath(pptx)
    pdf = os.path.join(os.path.abspath(outdir),
                       os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    app = _com("PowerPoint.Application")
    pres = None
    try:
        # Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0) — msoTrue/-1
        pres = app.Presentations.Open(pptx, -1, 0, 0)
        pres.SaveAs(pdf, 32)                       # 32 = ppSaveAsPDF
    finally:
        if pres is not None:
            pres.Close()
        try:
            if app.Presentations.Count == 0:       # чужие открытые деки не гасим
                app.Quit()
        except Exception:  # noqa: BLE001 — COM-зомби не должен ронять рендер
            pass
    if not os.path.isfile(pdf):
        return []
    pngs = pdf_to_png(pdf, outdir, dpi, only)
    try:
        os.remove(pdf)
    except OSError:
        pass
    return pngs


def find_soffice():
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    for c in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              "/opt/homebrew/bin/soffice", "/usr/local/bin/soffice",
              os.path.expandvars(r"%ProgramFiles%\LibreOffice\program\soffice.exe"),
              os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe")):
        if os.path.isfile(c):
            return c
    return None


def via_soffice(so, pptx, outdir, dpi, only=None):
    subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", outdir, pptx],
                   check=True, timeout=300, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    return pdf_to_png(pdf, outdir, dpi, only) if os.path.isfile(pdf) else []


def via_qlmanage(pptx, outdir):
    subprocess.run(["qlmanage", "-t", "-s", "1600", "-o", outdir, pptx],
                   check=True, timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    cover = os.path.join(outdir, os.path.basename(pptx) + ".png")
    if os.path.isfile(cover):
        dst = os.path.join(outdir, "slide-1.png")
        os.replace(cover, dst)
        sys.stderr.write("render: только обложка (qlmanage) — поставь PowerPoint/LibreOffice для всех слайдов\n")
        return [dst]
    return []


def contact_sheets(pngs, outdir, cols=4, rows=4, cell_w=640):
    """Контактный лист: сетка миниатюр с номерами слайдов, до cols×rows на лист
    (0.33.0, user-запрос «дек собирался час»: самопроверка B2–B13 по 1–2 листам
    вместо чтения каждого слайда отдельно; полноразмерно открываются только
    подозрительные)."""
    from PIL import Image, ImageDraw

    sheets, per = [], cols * rows
    for s0 in range(0, len(pngs), per):
        batch = pngs[s0:s0 + per]
        thumbs = []
        for p in batch:
            im = Image.open(p)
            th = im.resize((cell_w, int(im.height * cell_w / im.width)))
            thumbs.append(th)
        ch = max(t.height for t in thumbs)
        pad, cap = 8, 34
        sheet = Image.new("RGB", (cols * (cell_w + pad) + pad,
                                  min(rows, (len(batch) + cols - 1) // cols)
                                  * (ch + cap + pad) + pad), "#20242c")
        drw = ImageDraw.Draw(sheet)
        for i, th in enumerate(thumbs):
            x = pad + (i % cols) * (cell_w + pad)
            y = pad + (i // cols) * (ch + cap + pad)
            sheet.paste(th, (x, y + cap))
            n = int(_SLIDE_RE.search(os.path.basename(batch[i])).group(1))
            drw.text((x + 4, y + 6), f"слайд {n}", fill="#ffffff")
        out = os.path.join(outdir, f"contact-{s0 // per + 1}.png")
        sheet.save(out)
        sheets.append(out)
    return sheets


import re as _re_mod  # noqa: E402
_SLIDE_RE = _re_mod.compile(r"slide-0*(\d+)\.png$")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("-o", "--out")
    ap.add_argument("--slides", default=None,
                    help="печатать только эти слайды (1-based, напр. 4,11) — "
                         "точечная пере-проверка после правки")
    ap.add_argument("--contact-sheet", action="store_true",
                    help="сшить контактный лист (сетка миниатюр с номерами) и "
                         "печатать пути листов вместо отдельных PNG")
    a = ap.parse_args()
    pptx = os.path.abspath(a.pptx)
    if not os.path.isfile(pptx):
        sys.stderr.write(f"render: нет файла {pptx}\n")
        return 1
    outdir = a.out or (os.path.splitext(pptx)[0] + ".preview")
    os.makedirs(outdir, exist_ok=True)

    is_mac = platform.system() == "Darwin"
    is_win = platform.system() == "Windows"
    has_ppt = is_mac and os.path.isdir("/Applications/Microsoft PowerPoint.app")
    has_ppt_win = is_win and find_powerpnt_win()
    so = find_soffice()
    # --slides прокидывается в растеризацию: PDF движки отдают целиком
    # (выборочного экспорта нет), но PNG печём только выбранные (0.34.0)
    want = ({int(x) for x in a.slides.replace(" ", "").split(",") if x}
            if a.slides else None)

    cands = []
    if has_ppt:
        cands.append(("PowerPoint", lambda: via_powerpoint(pptx, outdir, a.dpi, want)))
    if has_ppt_win:
        cands.append(("PowerPoint COM",
                      lambda: via_powerpoint_com(pptx, outdir, a.dpi, want)))
    if so:
        cands.append(("LibreOffice", lambda: via_soffice(so, pptx, outdir, a.dpi, want)))
    if is_mac:
        cands.append(("qlmanage", lambda: via_qlmanage(pptx, outdir)))

    pngs = []
    for name, run in cands:
        try:
            pngs = run()
            if pngs:
                sys.stderr.write(f"render: [{name}] отрендерил {len(pngs)} слайд(ов)\n")
                break
            sys.stderr.write(f"render: [{name}] без превью, следующий движок...\n")
        except Exception as e:  # noqa: BLE001
            sys.stderr.write(f"render: [{name}] не справился ({e}); следующий движок...\n")

    if not pngs:
        sys.stderr.write("render: не удалось отрендерить. Поставь LibreOffice или PowerPoint.\n")
        return 1
    pngs = sorted(pngs)
    if want:
        # страховочный пост-фильтр (qlmanage want не понимает; повторный
        # прогон по outdir со старыми PNG)
        sel = [p for p in pngs
               if (m := _SLIDE_RE.search(os.path.basename(p)))
               and int(m.group(1)) in want]
        if sel:
            pngs = sel
        else:
            sys.stderr.write(f"render: --slides {a.slides}: таких слайдов нет "
                             f"(всего {len(pngs)}) — печатаю все\n")
    if a.contact_sheet:
        try:
            for p in contact_sheets(pngs, outdir):
                print(p)
            return 0
        except ImportError:
            sys.stderr.write("render: нет Pillow для контактного листа — "
                             "печатаю отдельные PNG\n")
    for p in pngs:
        print(p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
