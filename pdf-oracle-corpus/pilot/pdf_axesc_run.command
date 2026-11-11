#!/bin/bash
# pdf_axesc_run.command — LaunchServices fallback wrapper (queue-runner
# busy behind a lock-wait): cancel stacked Grant File Access panels via
# Escape, watchdog-bounded.
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_axesc_run.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== AXESC_RUN at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
TOK=$(cat "$HOME/word-oracle/WORD-LOCK.d/token" 2>/dev/null)
[ -n "$TOK" ] && bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 && echo HB_OK
osascript "$D/pdf_axesc.applescript" &
PID=$!
N=0
while kill -0 $PID 2>/dev/null; do
  if [ $N -ge 90 ]; then kill -9 $PID 2>/dev/null; echo "WATCHDOG_KILL osascript after 90s"; break; fi
  sleep 1; N=$((N+1))
done
wait $PID 2>/dev/null
echo "=== AXESC_RUN_DONE"
