#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_template.py — шаблон и шрифты скилла с публичного Я.Диска, по требованию.

Скилл распространяется БЕЗ тяжёлых .pptx (один шаблон ~27МБ — не лезет в лимиты
раздачи, а движок без шаблонов весит <1МБ). Шаблоны лежат на машинной полке —
публичной папке Я.Диска; листинг и скачивание — cloud-api.yandex.net БЕЗ OAuth
(тот же API, что в fetch_photo.py). Ревизия пинится по md5 в
Templates/manifest.json: движок качает ровно то, под что калиброваны layouts/.

Порядок резолва (build_deck зовёт resolve_template):
  1. канонический путь Templates/<tdir>/<имя>.pptx — дев-режим; md5 НЕ сверяем
     (файл правится дизайнером между пинами, семантику стережёт
     check_template_drift в build_deck);
  2. кэш ~/.cache/presentations-viscomms/<tdir>/<md5>/ — md5 сверен при записи,
     имя каталога = пин, бамп пина автоматически ведёт к новой закачке;
  3. скачивание с Диска → сверка md5+size против манифеста → атомарная запись
     в кэш; следом ставятся шрифты из fonts-ys.zip (идемпотентно, только
     недостающие файлы).

CLI:
  python3 scripts/fetch_template.py [--template еда|eda]   # печатает TEMPLATE=<путь>
  python3 scripts/fetch_template.py --check                # сверить пин с Диском, не качая
  python3 scripts/fetch_template.py --fonts                # только установить шрифты

Переопределение: PRESENTATIONS_VISCOMMS_CACHE (каталог кэша).

Фолбэк без сети: запроси файлы шаблона у команды (VISCOMMSSKILLS) и положи
в Templates/<tdir>/ — резолв возьмёт их шагом 1; шрифты из fonts-ys.zip
поставь двойным кликом или `--fonts`.

Публичные ссылки полок явно записаны в манифесте/brand-tokens. Скрипт выполняет
только HTTPS-загрузки через фиксированные хосты API и скачивания; пользовательские
файлы и содержимое презентации на Диск не отправляются.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
MANIFEST = os.path.join(SKILL_DIR, "Templates", "manifest.json")
API = "https://cloud-api.yandex.net/v1/disk/public/resources"
API_HOSTS = frozenset({"cloud-api.yandex.net"})
PUBLIC_LINK_HOSTS = frozenset({"disk.yandex.ru"})
DOWNLOAD_HOSTS = frozenset({"downloader.disk.yandex.ru"})
DOWNLOAD_DOMAINS = frozenset({"storage.yandex.net"})

TDIR_ALIASES = {"комитет": "committee", "комитет куда сходить": "committee",
                "комитет: куда сходить": "committee", "offline committee": "committee",
                "еда": "eda",
                "городские": "gor-services", "город": "gor-services",
                "городские сервисы": "gor-services", "gorod": "gor-services",
                "хурал": "ridetech-khural", "райдтех": "ridetech-khural", "райтех": "ridetech-khural",
                "хурал райдтеха": "ridetech-khural",
                "ridetech": "ridetech-khural", "khural": "ridetech-khural",
                "райдтех хурал": "ridetech-khural"}   # ЕДИНЫЙ источник алиасов tdir (build_deck/preflight/fetch_photo импортируют отсюда); user-facing ключ deckspec остаётся кириллическим


def _validate_https_url(url: str, allowed_hosts: frozenset[str], label: str,
                        allowed_domains: frozenset[str] = frozenset()) -> str:
    """Разрешить HTTPS URL только на точном хосте/домене из allowlist."""
    try:
        parsed = urllib.parse.urlsplit(url)
        port = parsed.port
    except (TypeError, ValueError) as error:
        raise ValueError(f"некорректный {label}: {error}") from error
    host = parsed.hostname or ""
    host_allowed = host in allowed_hosts or any(
        host == domain or host.endswith("." + domain)
        for domain in allowed_domains
    )
    if (parsed.scheme != "https" or not host_allowed
            or parsed.username is not None or parsed.password is not None
            or port not in (None, 443)):
        allowed = ", ".join(sorted(allowed_hosts | allowed_domains))
        raise ValueError(f"{label} разрешён только по HTTPS на хостах: {allowed}")
    return url


class _AllowlistedRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Не дать urllib уйти редиректом за пределы того же allowlist."""

    def __init__(self, allowed_hosts: frozenset[str], label: str,
                 allowed_domains: frozenset[str] = frozenset()):
        super().__init__()
        self.allowed_hosts = allowed_hosts
        self.allowed_domains = allowed_domains
        self.label = label

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        safe_url = _validate_https_url(newurl, self.allowed_hosts, self.label,
                                       self.allowed_domains)
        return super().redirect_request(req, fp, code, msg, headers, safe_url)


def _urlopen_allowlisted(url: str, allowed_hosts: frozenset[str],
                         label: str, timeout: int,
                         allowed_domains: frozenset[str] = frozenset()):
    safe_url = _validate_https_url(url, allowed_hosts, label, allowed_domains)
    opener = urllib.request.build_opener(
        _AllowlistedRedirectHandler(allowed_hosts, label, allowed_domains)
    )
    return opener.open(safe_url, timeout=timeout)


def validate_public_key(value: str) -> str:
    return _validate_https_url(value, PUBLIC_LINK_HOSTS, "публичная ссылка Диска")


def validate_download_href(value: str) -> str:
    return _validate_https_url(value, DOWNLOAD_HOSTS, "ссылка скачивания Диска",
                               DOWNLOAD_DOMAINS)


def load_manifest() -> dict:
    with open(MANIFEST, encoding="utf-8") as f:
        return json.load(f)


def public_key(man: dict, entry: dict | None = None) -> str:
    """Публичная ссылка: пер-шаблонная либо одна корневая из манифеста."""
    return validate_public_key((entry or {}).get("public_key") or man["public_key"])


def _remote_path(entry: dict, name: str) -> str:
    """Путь файла на полке. Пустой/отсутствующий remote_dir → файл в КОРНЕ
    публичной папки шаблона (пер-шаблонный public_key указывает прямо на неё)."""
    rd = (entry.get("remote_dir") or "").strip("/")
    return f"/{rd}/{name}" if rd else f"/{name}"


def cache_root() -> str:
    return (os.environ.get("PRESENTATIONS_VISCOMMS_CACHE")
            or os.path.join(os.path.expanduser("~"), ".cache", "presentations-viscomms"))


def _entry(man: dict, tdir: str) -> dict:
    tpl = man["templates"].get(tdir)
    if not tpl:
        raise SystemExit(f"в {MANIFEST} нет шаблона {tdir!r}; есть: {list(man['templates'])}")
    return tpl


def _md5(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _matches_pin(path: str, spec: dict) -> bool:
    return (os.path.getsize(path) == spec["size"]
            and _md5(path) == spec["md5"]
            and _sha256(path) == spec["sha256"])


def _require_pin(path: str, spec: dict, label: str) -> str:
    if _matches_pin(path, spec):
        return path
    if os.environ.get("PRESENTATIONS_VISCOMMS_ALLOW_UNPINNED_TEMPLATE") == "1":
        print(f"  ⚠ {label} не совпадает с SHA-256 пином; разрешено явным dev opt-in")
        return path
    raise SystemExit(
        f"{label} не совпадает с пином manifest.json (size/md5/SHA-256). "
        "Замени файл на закреплённую ревизию или для осознанной разработки задай "
        "PRESENTATIONS_VISCOMMS_ALLOW_UNPINNED_TEMPLATE=1.")


def _api_json(url: str) -> dict:
    last = None
    for _ in range(3):
        try:
            with _urlopen_allowlisted(url, API_HOSTS, "Yandex Disk API", 60) as r:
                return json.load(r)
        except Exception as e:  # noqa: BLE001 — сетевые флапы, ретраим
            last = e
            time.sleep(2)
    raise last


def _download_href(pk: str, remote_path: str) -> str:
    q = urllib.parse.urlencode({"public_key": pk, "path": remote_path})
    return validate_download_href(_api_json(f"{API}/download?{q}")["href"])


def _remote_meta(pk: str, remote_path: str) -> dict:
    q = urllib.parse.urlencode({"public_key": pk, "path": remote_path})
    return _api_json(f"{API}?{q}")


def _fetch_file(pk: str, remote_path: str, dest: str, spec: dict) -> None:
    """Скачивание с проверкой SHA-256+md5+size против манифеста,
    иначе на Диске лежит НЕ та ревизия, под которую калиброван движок."""
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = tempfile.NamedTemporaryFile(prefix=".download-", dir=os.path.dirname(dest), delete=False)
    part = tmp.name
    tmp.close()
    href = _download_href(pk, remote_path)
    with _urlopen_allowlisted(href, DOWNLOAD_HOSTS,
                              "ссылка скачивания Диска", 300,
                              DOWNLOAD_DOMAINS) as r, \
            open(part, "wb") as f:
        shutil.copyfileobj(r, f)
    got_size, got_md5, got_sha256 = os.path.getsize(part), _md5(part), _sha256(part)
    if (got_md5 != spec["md5"] or got_size != spec["size"]
            or got_sha256 != spec["sha256"]):
        os.remove(part)
        raise SystemExit(
            f"пин не сошёлся для {remote_path!r}: md5={got_md5} sha256={got_sha256} size={got_size}; "
            f"манифест ждёт md5={spec['md5']} sha256={spec['sha256']} size={spec['size']}. "
            f"перекалибровки — нужен бамп пина в {MANIFEST} ВМЕСТЕ с layouts/ "
            f"(inventory --diff, калибровка), см. references/quality-gates.md")
    os.replace(part, dest)


def download_file(href: str, dest: str, timeout: int = 300) -> None:
    """Скачать публичный файл через фиксированный download-хост, атомарно."""
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = tempfile.NamedTemporaryFile(prefix=".download-", dir=os.path.dirname(dest), delete=False)
    part = tmp.name
    tmp.close()
    try:
        with _urlopen_allowlisted(href, DOWNLOAD_HOSTS,
                                  "ссылка скачивания Диска", timeout,
                                  DOWNLOAD_DOMAINS) as src, \
                open(part, "wb") as out:
            shutil.copyfileobj(src, out)
        os.replace(part, dest)
    finally:
        if os.path.exists(part):
            os.remove(part)


# --- шрифты -----------------------------------------------------------------

def _font_dest() -> str:
    system = platform.system()
    if system == "Darwin":
        return os.path.expanduser("~/Library/Fonts")
    if system == "Windows":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        return os.path.join(base, "Microsoft", "Windows", "Fonts")
    return os.path.expanduser("~/.local/share/fonts")


def install_fonts_from_zip(zip_path: str) -> list[str]:
    """Идемпотентно ставит .ttf/.otf из архива в пользовательский каталог шрифтов
    (по basename, плоско — архив бывает с папкой в легаси-кодировке).
    Возвращает список реально установленных (новых) файлов."""
    dest = _font_dest()
    os.makedirs(dest, exist_ok=True)
    installed = []
    with zipfile.ZipFile(zip_path) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            base = info.filename.replace("\\", "/").rsplit("/", 1)[-1]
            if not base.lower().endswith((".ttf", ".otf")):
                continue
            target = os.path.join(dest, base)
            if os.path.exists(target):
                continue
            with z.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            installed.append(base)
    if platform.system() == "Linux" and shutil.which("fc-cache"):
        subprocess.run(["fc-cache", "-f", dest],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    return installed


def _font_globs(tdir: str) -> list[str]:
    """Glob-паттерны имён шрифтов шаблона — префиксы из custom-profile.json
    (fonts.files), фолбэк «YS *» (бренд Еды). Без хардкода семейства."""
    prof = os.path.join(SKILL_DIR, "layouts", tdir, "custom-profile.json")
    try:
        with open(prof, encoding="utf-8") as f:
            files = json.load(f)["fonts"]["files"]
        return [p + "*" for p in sorted({v.split("-")[0].strip()
                                         for v in files.values()})] or ["YS *"]
    except (OSError, ValueError, KeyError):
        return ["YS *"]


def _win_fonts_present(tdir: str) -> bool:
    """Шрифты шаблона на Windows: пользовательская установка или системная папка."""
    import glob
    dirs = [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
            r"C:\Windows\Fonts"]
    return any(glob.glob(os.path.join(d, pat))
               for d in dirs if d and os.path.isdir(d) for pat in _font_globs(tdir))


def ensure_fonts(tdir: str) -> None:
    """Шрифты нужны и PowerPoint-рендеру, и глиф-замеру lib/fit.py (ищет в
    ~/Library/Fonts). Берём fonts-ys.zip: локальный Templates/ → кэш → Диск.

    Windows: НЕ копируем — файл в папке Fonts без записи в реестре системой
    не активируется (молчаливая подмена типографики). Только руками через
    проводник: ОС сама регистрирует шрифт."""
    man = load_manifest()
    entry = _entry(man, tdir)
    fz = entry["fonts"]
    if "system_required" in fz:
        missing = [name for name in fz["system_required"]
                   if not os.path.isfile(os.path.join(_font_dest(), name))]
        if missing:
            raise SystemExit("Не установлены шрифты локального шаблона: " + ", ".join(missing))
        return
    if platform.system() == "Windows":
        if _win_fonts_present(tdir):
            return
        print(f"  Windows: шрифты YS не установлены — автоустановка здесь не работает.\n"
              f"  Скачай {fz['name']} с полки {public_key(man, entry)} , распакуй,\n"
              "  выдели все файлы шрифтов → ПКМ → «Установить». Без этого PowerPoint\n"
              "  молча подменит типографику в собранном деке.")
        return
    local = os.path.join(SKILL_DIR, "Templates", tdir, fz["name"])
    cached = os.path.join(cache_root(), tdir, fz["sha256"], fz["name"])
    zip_path = local if os.path.exists(local) else cached
    if os.path.exists(zip_path):
        _require_pin(zip_path, fz, "архив шрифтов")
    if not os.path.exists(zip_path):
        pk = public_key(man, entry)
        print(f"  качаю шрифты {fz['name']} с Я.Диска…")
        _fetch_file(pk, _remote_path(entry, fz["name"]), cached, fz)
        zip_path = cached
    new = install_fonts_from_zip(zip_path)
    if new:
        print(f"  шрифты установлены ({len(new)}): {', '.join(sorted(new))}")


# --- резолв шаблона ----------------------------------------------------------

def resolve_template(tdir: str) -> str:
    """Абсолютный путь к .pptx шаблона. Локальный Templates/ (дев) → кэш → Диск."""
    tdir = TDIR_ALIASES.get(tdir, tdir)
    man = load_manifest()
    entry = _entry(man, tdir)
    tpl = entry["template"]

    local = os.path.join(SKILL_DIR, "Templates", tdir, tpl["name"])
    if os.path.exists(local):
        return _require_pin(local, tpl, "локальный шаблон")

    # имя не совпало — ищем ревизию по md5 среди любых *.pptx в папке шаблона
    # (фидбек G, Linux/облако 2026-07-16: локальный файл под другим именем
    # раньше отправлял резолв в сеть, а сети в облаке нет)
    tpl_dir = os.path.join(SKILL_DIR, "Templates", tdir)
    if os.path.isdir(tpl_dir):
        for cand in sorted(os.listdir(tpl_dir)):
            path = os.path.join(tpl_dir, cand)
            if not cand.lower().endswith(".pptx") or not os.path.isfile(path):
                continue
            if _matches_pin(path, tpl):
                print(f"  шаблон найден по md5-пину: {cand} "
                      f"(имя ≠ манифестному {tpl['name']} — это ок)")
                return path

    cached = os.path.join(cache_root(), tdir, tpl["sha256"], tpl["name"])
    if os.path.exists(cached):
        return _require_pin(cached, tpl, "шаблон в кэше")

    if entry.get("local_only"):
        raise SystemExit(
            f"Локальный шаблон {tdir!r} отсутствует: {local}. "
            "Восстанови файл из установленного набора; этот шаблон не скачивается с публичной полки.")

    pk = public_key(man, entry)
    print(f"  шаблона нет локально — качаю {tpl['name']} (~{tpl['size'] >> 20}МБ) с Я.Диска…")
    _fetch_file(pk, _remote_path(entry, tpl["name"]), cached, tpl)
    print(f"  шаблон в кэше: {cached}")
    ensure_fonts(tdir)
    return cached


def check_pin(tdir: str) -> int:
    """Сверка пина с текущим состоянием Диска (md5 из метаданных, без закачки)."""
    man = load_manifest()
    entry = _entry(man, tdir)
    if entry.get("local_only"):
        path = resolve_template(tdir)
        _require_pin(path, entry["template"], "локальный шаблон")
        print(f"  {entry['template']['name']}: локальный size/md5/SHA-256 пин совпадает")
        return 0
    pk = public_key(man, entry)
    ok = True
    for key in ("template", "fonts"):
        f = entry[key]
        meta = _remote_meta(pk, _remote_path(entry, f["name"]))
        match = meta.get("md5") == f["md5"] and meta.get("size") == f["size"]
        ok &= match
        print(f"  {f['name']}: {'✓ пин совпадает' if match else '✗ РАЗОШЁЛСЯ'} "
              f"(диск md5={meta.get('md5')} size={meta.get('size')}, "
              f"пин md5={f['md5']} size={f['size']})")
    if not ok:
        print("  полку обновили без бампа пина: перекалибровать layouts/ и обновить manifest.json")
    return 0 if ok else 1


def main() -> int:
    ap = argparse.ArgumentParser(description="шаблон и шрифты скилла с Я.Диска по пину")
    ap.add_argument("--template", default="eda", help="tdir шаблона (еда|eda)")
    ap.add_argument("--check", action="store_true", help="сверить пин с Диском, не качая")
    ap.add_argument("--fonts", action="store_true", help="только установить шрифты")
    a = ap.parse_args()
    tdir = TDIR_ALIASES.get(a.template.lower(), a.template.lower())
    if a.check:
        return check_pin(tdir)
    if a.fonts:
        ensure_fonts(tdir)
        return 0
    print(f"TEMPLATE={resolve_template(tdir)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
