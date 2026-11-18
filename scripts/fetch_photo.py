#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_photo.py — фото на замену из одобренного пула шаблона (Я.Диск), по требованию.

Пул — ПУБЛИЧНАЯ папка Яндекс Диска; листинг и скачивание идут через
cloud-api.yandex.net БЕЗ OAuth. Фото в репо/скилле не храним: пул одного шаблона
~62МБ, на много шаблонов — гигабайты (решение user 2026-07-02). URL пула живёт
в layouts/<шаблон>/brand-tokens.json → "photo_pool".

Использование:
  python3 scripts/fetch_photo.py list  [--template еда]
  python3 scripts/fetch_photo.py get Бьюти_01.jpg [Дети_02.jpg ...] [--template еда]
  python3 scripts/fetch_photo.py get-all [--template еда]

Скачанное кешируется в приватном user-cache `~/.cache/presentations-viscomms/photos/`
(каталог создаётся с mode 0700,
оттуда же удобно вставлять и рендерить); повторный get отдаёт кеш без сети.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import urllib.parse

import fetch_template   # единый источник TDIR_ALIASES (scripts/ уже на sys.path: прямой запуск и импорт из build_deck)

REPO = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
API = "https://cloud-api.yandex.net/v1/disk/public/resources"
CACHE_ROOT = os.path.join(fetch_template.cache_root(), "photo-pool")


def _safe_component(value: str, label: str) -> str:
    if value in {"", ".", ".."} or os.path.basename(value) != value or not re.fullmatch(r"[\w .@()+-]+", value):
        raise ValueError(f"unsafe {label}: {value!r}")
    return value


def pool_url(template: str) -> str:
    p = os.path.join(REPO, "layouts", template, "brand-tokens.json")
    with open(p, encoding="utf-8") as f:
        tokens = json.load(f)
    pool = tokens.get("photo_pool") or {}
    if not pool.get("url"):
        raise SystemExit(f"в {p} нет photo_pool.url — у шаблона не заведён пул фото")
    return fetch_template.validate_public_key(pool["url"])


def _api(endpoint: str, **params) -> dict:
    qs = urllib.parse.urlencode(params)
    return fetch_template._api_json(f"{API}{endpoint}?{qs}")


def list_pool(template: str) -> list[dict]:
    """Все файлы пула: [{name, size, mime_type, path}], с пагинацией."""
    key, items, offset = pool_url(template), [], 0
    while True:
        d = _api("", public_key=key, limit=200, offset=offset)
        emb = d.get("_embedded", {})
        batch = emb.get("items", [])
        items += [i for i in batch if i.get("type") == "file"]
        offset += len(batch)
        if offset >= emb.get("total", 0) or not batch:
            return items


def fetch(template: str, name: str) -> str:
    template = _safe_component(template, "template")
    name = _safe_component(name, "photo name")
    """Скачать одно фото (или отдать из кеша). Возвращает локальный путь."""
    dst_dir = os.path.join(CACHE_ROOT, template)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, name)
    if os.path.islink(dst):
        raise ValueError(f"refusing symlink cache target: {dst}")
    if os.path.exists(dst) and os.path.getsize(dst) > 0:
        return dst
    d = _api("/download", public_key=pool_url(template), path=f"/{name}")
    fetch_template.download_file(d["href"], dst)
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["list", "get", "get-all"])
    ap.add_argument("names", nargs="*")
    ap.add_argument("--template", default="eda")
    a = ap.parse_args()
    a.template = fetch_template.TDIR_ALIASES.get(a.template.lower(), a.template.lower())

    if a.cmd == "list":
        for it in list_pool(a.template):
            print(f"{it['name']:24} {it.get('size', 0) // 1024:>6} KB")
        return
    names = a.names if a.cmd == "get" else [i["name"] for i in list_pool(a.template)]
    if not names:
        raise SystemExit("get: укажи имена файлов (см. list)")
    for n in names:
        print(fetch(a.template, n))


if __name__ == "__main__":
    main()
