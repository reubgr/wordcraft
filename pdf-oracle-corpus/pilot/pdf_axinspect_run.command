#!/bin/bash
# pdf_axinspect_run.command — LaunchServices fallback wrapper for the
# read-only AX census (queue-runner busy behind a lock-wait).
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_axinspect_run.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== AXINSPECT_RUN at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
osascript "$D/pdf_axinspect.applescript" 2>&1
echo "=== AXINSPECT_RUN_DONE"
