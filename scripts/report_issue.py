#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Подготовить воспроизводимый багрепорт для очереди VISCOMMSSKILLS.

Скрипт работает только локально: собирает версию скилла, пин шаблона,
платформу и пути к вложениям. Сетевых запросов и авторизации здесь нет.

После явного подтверждения пользователя агент создаёт тикет через официальный
Tracker MCP и прикладывает перечисленные файлы. Если Tracker MCP недоступен,
скрипт также печатает ссылку на веб-форму с предзаполненными полями.

Пример:
  python3 scripts/report_issue.py --summary "timeline молчит при N=4" \
      --expected "честная ошибка с ёмкостями" --got "слайд собрался пустым" \
      --deckspec deckspec.json --log build.log --render slide_07.png \
      --template gor-services

Баг (дефолт) — тип task + тег `bug`; идея/доработка (`--kind improvement`) —
тип improvement. Отдельного типа bug в очереди нет — различаем тегом.
"""
from __future__ import annotations

import argparse
import json
import os
import platform
import sys
import time
import urllib.parse

REPO = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
QUEUE = "VISCOMMSSKILLS"
WEB = "https://st.yandex-team.ru"
URL_DESC_LIMIT = 1000


def skill_version() -> str:
    try:
        with open(os.path.join(REPO, "SKILL.md"), encoding="utf-8") as f:
            for line in f:
                if line.startswith("version:"):
                    return line.split(":", 1)[1].strip()
    except OSError:
        pass
    return "?"


def template_pin(tkey: str | None) -> str:
    """'gor-services (template-gs-v2.0.pptx, пин 72e6d578)' либо оба."""
    try:
        with open(os.path.join(REPO, "Templates", "manifest.json"),
                  encoding="utf-8") as f:
            man = json.load(f)
    except OSError:
        return tkey or "?"
    keys = [tkey] if tkey else sorted(man.get("templates", {}))
    parts = []
    for key in keys:
        template = (man.get("templates", {}).get(key) or {}).get("template") or {}
        if template:
            parts.append(
                f"{key} ({template.get('name', '?')}, "
                f"пин {str(template.get('md5', ''))[:8]})"
            )
        else:
            parts.append(key)
    return "; ".join(parts) if parts else "?"


def build_description(args: argparse.Namespace) -> str:
    lines = [
        f"**Что ожидали:** {args.expected or '(заполните)'}",
        f"**Что получили:** {args.got or '(заполните)'}",
        "",
        "**Контекст** (собран `scripts/report_issue.py`):",
        f"- скилл: presentation-assembler v{skill_version()}",
        f"- шаблон: {template_pin(args.template)}",
        f"- платформа: {platform.platform()}, python {platform.python_version()}",
        f"- дата: {time.strftime('%Y-%m-%d %H:%M %Z')}",
    ]
    attached = [path for path in [args.deckspec, args.log, *(args.render or [])]
                if path]
    if attached:
        names = ", ".join(os.path.basename(path) for path in attached)
        lines += ["", f"**Файлы** (deckspec / вывод сборки / рендер): {names}"]
    return "\n".join(lines)


def existing_files(args: argparse.Namespace) -> list[str]:
    files = []
    for path in [args.deckspec, args.log, *(args.render or [])]:
        if not path:
            continue
        if os.path.isfile(path):
            files.append(os.path.abspath(path))
        else:
            print(f"⚠ файла нет, пропускаю: {path}", file=sys.stderr)
    return files


def createticket_url(summary: str, description: str,
                     issue_type: str, tags: list[str]) -> str:
    query = {
        "queue": QUEUE,
        "summary": summary,
        "description": description[:URL_DESC_LIMIT],
        "type": issue_type,
        "tags": ",".join(tags),
    }
    return f"{WEB}/createTicket?" + urllib.parse.urlencode(
        {key: value for key, value in query.items() if value}
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Подготовить тикет на доработку скилла в очередь " + QUEUE
    )
    parser.add_argument("--summary", required=True, help="заголовок тикета")
    parser.add_argument("--expected", help="что ожидали")
    parser.add_argument("--got", help="что получили")
    parser.add_argument("--deckspec", help="путь к deckspec.json или промпту")
    parser.add_argument("--log", help="путь к полному выводу сборки")
    parser.add_argument("--render", nargs="*", help="PNG проблемного слайда")
    parser.add_argument("--template", choices=["eda", "gor-services"],
                        help="какой шаблон затронут")
    parser.add_argument("--kind", choices=["bug", "improvement"], default="bug")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="совместимость со старым CLI; скрипт в любом случае ничего не отправляет",
    )
    args = parser.parse_args()

    issue_type = "task" if args.kind == "bug" else "improvement"
    tags = (["bug"] if args.kind == "bug" else [])
    if args.template:
        tags.append(args.template)
    description = build_description(args)
    files = existing_files(args)

    print("Отчёт подготовлен локально; сетевых запросов не выполнялось.\n")
    print("После подтверждения пользователя агент должен создать тикет через "
          "Tracker MCP со следующими полями:\n")
    print(f"  queue: {QUEUE}")
    print(f"  summary: {args.summary}")
    print(f"  type: {issue_type}")
    print(f"  tags: {', '.join(tags) if tags else '(нет)'}")
    print("\ndescription:\n")
    print(description)
    if files:
        print("\nattachments:")
        for path in files:
            print("  " + path)
    print("\nЕсли Tracker MCP недоступен, откройте веб-форму:\n")
    print("  " + createticket_url(args.summary, description, issue_type, tags))
    return 0


if __name__ == "__main__":
    sys.exit(main())
