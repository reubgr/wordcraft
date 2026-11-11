#!/bin/bash
# pdf_opts5_probe.command — gate probe for the v5 resumption (#6153):
# compile-check happens VM-side; this job takes the WORD-LOCK, runs ONE
# equality-based options snapshot via the proven `settings` route
# (pdf_opts5.applescript), releases, and exits. No sets, no documents.
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_opts5_probe.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== OPTS5_PROBE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
runwd() {
  local limit=$1; shift
  "$@" &
  local pid=$! n=0
  while kill -0 $pid 2>/dev/null; do
    if [ $n -ge $limit ]; then kill -9 $pid 2>/dev/null; wait $pid 2>/dev/null; echo "WATCHDOG_KILL after ${limit}s: $*"; return 124; fi
    sleep 1; n=$((n+1))
  done
  wait $pid; return $?
}
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-opts5-probe --wait 1200 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== OPTS5_PROBE_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
runwd 90 osascript "$D/pdf_opts5.applescript" snapshot 2>&1
echo "SNAPSHOT_RC=$?"
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "DOCS " & (count of documents)
end timeout' 2>&1
echo "=== OPTS5_PROBE_DONE"
