#!/bin/bash
# pdf_grantcancel_run.command — LaunchServices fallback wrapper,
# watchdog-bounded (75s): cancel stacked Grant File Access panels.
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_grantcancel_run.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== GRANTCANCEL_RUN at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
osascript "$D/pdf_grantcancel.applescript" &
PID=$!
N=0
while kill -0 $PID 2>/dev/null; do
  if [ $N -ge 75 ]; then kill -9 $PID 2>/dev/null; echo "WATCHDOG_KILL osascript after 75s"; break; fi
  sleep 1; N=$((N+1))
done
wait $PID 2>/dev/null
echo "=== GRANTCANCEL_RUN_DONE"
