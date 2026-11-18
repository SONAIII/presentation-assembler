#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_asset.py — ассеты (логотипы/флаги/иконки) с полки Я.Диска по имени.

Полка — та же публичная папка, что у шаблонов (`Templates/manifest.json` →
public_key), подпапка `assets/<категория>/`. Категории (например logos,
flags, icons) создаёт и наполняет команда скилла — конвертером
library_from_pptx.py из дизайнерских pptx-библиотек. Пины не нужны — как у фото-пула,
кеш по имени: `~/.cache/presentations-viscomms/assets/<категория>/`.

  python3 scripts/fetch_asset.py list logos
  python3 scripts/fetch_asset.py get logos "Такси.png" "Алиса.png"
  python3 scripts/fetch_asset.py get-all flags
"""
from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.parse

import fetch_template   # единый безопасный клиент публичного Диска

REPO = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
API = "https://cloud-api.yandex.net/v1/disk/public/resources"
CACHE_ROOT = os.path.join(fetch_template.cache_root(), "asset-cache")


def _safe_component(value: str, label: str) -> str:
    if value in {"", ".", ".."} or os.path.basename(value) != value or not re.fullmatch(r"[\w .@()+-]+", value):
        raise ValueError(f"unsafe {label}: {value!r}")
    return value


def shelf_key() -> str:
    with open(os.path.join(REPO, "Templates", "manifest.json"),
              encoding="utf-8") as f:
        return fetch_template.validate_public_key(json.load(f)["public_key"])


def _api(endpoint: str, **params) -> dict:
    qs = urllib.parse.urlencode(params)
    try:
        return fetch_template._api_json(f"{API}{endpoint}?{qs}")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise SystemExit(
                "на полке нет такого пути — категория ещё не выложена? "
                "(assets/<категория>/ создаёт владелец полки, "
                "категорию создаёт команда скилла — тикет в VISCOMMSSKILLS)") from e
        raise


def list_assets(category: str) -> list[dict]:
    category = _safe_component(category, "category")
    key, items, offset = shelf_key(), [], 0
    while True:
        d = _api("", public_key=key, path=f"/assets/{category}",
                 limit=200, offset=offset)
        emb = d.get("_embedded", {})
        batch = emb.get("items", [])
        items += [i for i in batch if i.get("type") == "file"]
        offset += len(batch)
        if offset >= emb.get("total", 0) or not batch:
            return items


def fetch(category: str, name: str) -> str:
    category = _safe_component(category, "category")
    name = _safe_component(name, "asset name")
    dst_dir = os.path.join(CACHE_ROOT, category)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, name)
    if os.path.islink(dst):
        raise ValueError(f"refusing symlink cache target: {dst}")
    if os.path.exists(dst) and os.path.getsize(dst) > 0:
        return dst
    d = _api("/download", public_key=shelf_key(),
             path=f"/assets/{category}/{name}")
    fetch_template.download_file(d["href"], dst)
    return dst


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["list", "get", "get-all"])
    ap.add_argument("category", help="logos | flags | icons | …")
    ap.add_argument("names", nargs="*")
    a = ap.parse_args()

    if a.cmd == "list":
        for it in list_assets(a.category):
            print(f"{it['name']:40} {it.get('size', 0) // 1024:>6} KB")
        return
    names = a.names if a.cmd == "get" else [i["name"] for i in list_assets(a.category)]
    if not names:
        raise SystemExit("get: укажи имена файлов (см. list)")
    for n in names:
        print(fetch(a.category, n))


if __name__ == "__main__":
    main()
