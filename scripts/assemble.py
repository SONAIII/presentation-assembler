#!/usr/bin/env python3
"""Presentation Assembler CLI — единая точка входа.

Здесь живёт маршрутизация source profiles и сборки PPTX. Данные извлекает
доступный пользователю connector; CLI не содержит provider-specific adapters.

Бизнес-логика данных и PPTX остаётся в профильных модулях. Платформенные
обёртки не должны дублировать эту маршрутизацию.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Manage sources or build a PPTX")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("templates", help="List selectable templates without network access")

    sources = sub.add_parser("sources", help="Validate or inspect workspace source profiles")
    sources.add_argument("source_args", nargs=argparse.REMAINDER,
                         help="arguments for source_profiles.py")

    build = sub.add_parser("build", help="Build an approved deckspec into PPTX")
    build.add_argument("deckspec")
    build.add_argument("-o", "--out", required=True)
    build.add_argument("--dry-run", action="store_true")

    args = parser.parse_args()
    scripts = Path(__file__).resolve().parent

    if args.command == "templates":
        manifest = json.loads((scripts.parent / "Templates/manifest.json").read_text())
        labels = {"eda": "Еда", "gor-services": "Городские сервисы", "ridetech-khural": "Хурал Райдтеха"}
        print(json.dumps([{"id": name, "name": entry.get("display_name", labels.get(name, name)),
                           "local_only": bool(entry.get("local_only"))}
                          for name, entry in manifest["templates"].items()], ensure_ascii=False, indent=2))
        return 0

    if args.command == "sources":
        cmd = [sys.executable, str(scripts / "source_profiles.py"),
               *args.source_args]
    else:
        cmd = [sys.executable, str(scripts / "build_deck.py"),
               args.deckspec, "-o", args.out]
        if args.dry_run:
            cmd.append("--dry-run")

    return subprocess.run(cmd, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
