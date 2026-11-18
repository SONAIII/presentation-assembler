#!/bin/sh
# bootstrap.sh — онбординг окружения С НУЛЯ, единственный скрипт скилла без
# python-зависимостей (курица-яйцо: preflight сам на python).
#
#   bash scripts/bootstrap.sh
#
# Делает по порядку: 1) выбирает Python >=3.11; 2) создаёт изолированный venv
# в кэше; 3) ставит точный lock; 4) запускает полный preflight.
#
# Windows не покрывает (шелла может не быть) — см. SKILL.md «Первый запуск».

set -u
SKILL_DIR=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)

echo "== bootstrap: окружение presentation-assembler =="

# --- 1. управляемый Python --------------------------------------------------
PYTHON_BIN=${PRESENTATION_ASSEMBLER_PYTHON:-${STEFANIA_PYTHON:-}}
if [ -z "$PYTHON_BIN" ]; then
    BUNDLED="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
    if [ -x "$BUNDLED" ]; then PYTHON_BIN=$BUNDLED; else PYTHON_BIN=$(command -v python3 || true); fi
fi
if [ -z "$PYTHON_BIN" ] || ! "$PYTHON_BIN" -c 'import sys; raise SystemExit(sys.version_info < (3,11))'; then
    echo "  ✗ нужен Python >=3.11. Укажи PRESENTATION_ASSEMBLER_PYTHON=/path/to/python3"
    exit 1
fi
echo "  ✓ runtime — $($PYTHON_BIN -V 2>&1)"

# --- 2. изолированный venv --------------------------------------------------
CACHE_ROOT=${PRESENTATION_ASSEMBLER_CACHE:-"$HOME/.cache/presentation-assembler"}
VENV="$CACHE_ROOT/runtime/py311"
if [ ! -x "$VENV/bin/python" ]; then "$PYTHON_BIN" -m venv "$VENV" || exit 1; fi
PIP_INDEX=${PRESENTATION_ASSEMBLER_PIP_INDEX:-https://pypi.yandex-team.ru/simple/}
echo "  ставлю lock в $VENV …"
"$VENV/bin/python" -m pip install --quiet --index-url "$PIP_INDEX" -r "$SKILL_DIR/requirements.txt" \
    || { echo "  ✗ установка пакетов не удалась; проверь доступ к package index"; exit 1; }
PUBLIC_INDEX=${PRESENTATION_ASSEMBLER_PUBLIC_PIP_INDEX:-https://pypi.org/simple/}
"$VENV/bin/python" -m pip install --quiet --index-url "$PUBLIC_INDEX" --only-binary=:all: \
    --no-deps --require-hashes -r "$SKILL_DIR/requirements-public.txt" \
    || { echo "  ✗ hash-locked uharfbuzz не установился"; exit 1; }
echo "  ✓ пакеты движка"

# --- 3. полный диагноз ------------------------------------------------------
echo "  runtime: $VENV/bin/python"
exec "$VENV/bin/python" "$SKILL_DIR/scripts/preflight.py"
