#!/bin/bash
# pdf_grantdump_run.command — LaunchServices fallback wrapper,
# watchdog-bounded (60s).
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_grantdump_run.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== GRANTDUMP_RUN at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
osascript "$D/pdf_grantdump.applescript" &
PID=$!
N=0
while kill -0 $PID 2>/dev/null; do
  if [ $N -ge 60 ]; then kill -9 $PID 2>/dev/null; echo "WATCHDOG_KILL osascript after 60s"; break; fi
  sleep 1; N=$((N+1))
done
wait $PID 2>/dev/null
echo "=== GRANTDUMP_RUN_DONE"
