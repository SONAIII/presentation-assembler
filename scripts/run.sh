#!/bin/sh
set -eu
SKILL_DIR=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
CACHE_ROOT=${PRESENTATION_ASSEMBLER_CACHE:-"$HOME/.cache/presentation-assembler"}
PY="$CACHE_ROOT/runtime/py311/bin/python"
if [ ! -x "$PY" ]; then
    echo "runtime missing; run: sh $SKILL_DIR/scripts/bootstrap.sh" >&2
    exit 2
fi
if [ "$#" -eq 0 ]; then
    echo "usage: sh scripts/run.sh scripts/<command>.py [args...]" >&2
    exit 2
fi
exec "$PY" "$@"
