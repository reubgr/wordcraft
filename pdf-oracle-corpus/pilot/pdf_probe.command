#!/bin/bash
# pdf_probe.command — lock-holding recovery + options probe.
# Acquires the WORD-LOCK itself (bounded wait); only under the lock does
# it close documents (autorecovery/staged copies left by the crashed
# Word) and probe Track Changes Options behaviour. Releases on exit.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
OUT=$D/pdf_probe.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_PROBE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-corpus-pilot-probe --wait 900 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  echo "=== PDF_PROBE_DONE rc=lock"
  exit 3
fi
echo "LOCK_ACQUIRED"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
rm -f "$C/pdf_probe_log.txt"
osascript "$D/pdf_probe.applescript" 2>&1
echo "--- probe log:"
cat "$C/pdf_probe_log.txt" 2>/dev/null
echo "=== PDF_PROBE_DONE"
