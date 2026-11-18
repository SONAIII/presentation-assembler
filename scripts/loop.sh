#!/usr/bin/env bash
# VENDORED (locally adapted) — source of truth: pptx-crash-check skill
#   (~/.claude/skills/pptx-crash-check/scripts/loop.sh).
# Kept in sync EXCEPT the assembly-fix reference: upstream says backend/app.py;
# here it points to lib/clone.py + the chart-guard in scripts/build_deck.py.
# Re-vendor structural changes from upstream, but preserve that local reference.
# Self-sustaining diagnose loop.
#
# Usage: loop.sh <path-to-pptx> [<path-to-known-good.pptx>]
#
# 1. Runs check_pptx.py against the file. If anything FAILs, prints results and exits 1
#    — you fix that and re-run.
# 2. If everything PASSes, opens the file in PowerPoint so the human can confirm.
# 3. If the human confirms it opens clean: done.
# 4. If the human reports Repair: runs deep_diagnose.py against the optional good file
#    to surface candidate failure modes that aren't catalogued yet.
#
# Anything found in step 4 gets manually added to known_issues.md + check_pptx.py
# (and the underlying assembly fix — in this skill: lib/clone.py or the
# chart-guard in scripts/build_deck.py) before the next iteration.
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ $# -lt 1 || $# -gt 2 ]]; then
    echo "Usage: $0 <bad.pptx> [<good.pptx>]" >&2
    exit 2
fi
BAD="$1"
GOOD="${2:-}"

if [[ ! -f "$BAD" ]]; then echo "File not found: $BAD" >&2; exit 2; fi

echo "==> Static checks on $BAD"
if ! python3 "$SCRIPT_DIR/check_pptx.py" "$BAD"; then
    echo
    echo "Static checks failed. Fix those first, then re-run loop.sh."
    exit 1
fi

echo
echo "==> All known checks pass. Opening in PowerPoint for human verification."
open -a "Microsoft PowerPoint" "$BAD"
echo
echo "When PowerPoint finishes loading, report back:"
echo "  - 'ok'      → file opens clean, we're done"
echo "  - 'repair'  → Repair dialog appeared, this is a new failure mode"
echo
if [[ -n "$GOOD" && -f "$GOOD" ]]; then
    echo "If you report 'repair', the next step is:"
    echo "  python3 $SCRIPT_DIR/deep_diagnose.py \"$BAD\" \"$GOOD\""
    echo "then catalogue the finding in known_issues.md and add a check."
fi
