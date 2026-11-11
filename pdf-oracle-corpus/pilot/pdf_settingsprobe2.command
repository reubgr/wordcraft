#!/bin/bash
# pdf_settingsprobe2.command — probe 3b (#6153): file-based rerun of the
# settings-formulation + save-target tests (probe 3's heredoc osascript
# calls silently no-oped in the runner context). Lock-held, self-restoring.
D=/Users/muse/compare-vs-word/dialog
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_settingsprobe2.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== SETTINGSPROBE2 at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-settingsprobe2 --wait 900 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== SETTINGSPROBE2_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
echo "--- settings formulations"
runwd 75 osascript "$D/pdf_settings_form.applescript" 2>&1
echo "--- save-target test"
mkdir -p "$P/out/tc-p2m-multiauthor"
runwd 135 osascript "$D/pdf_save_test.applescript" 2>&1
ls -la "$P/out/tc-p2m-multiauthor/" 2>&1
echo "=== SETTINGSPROBE2_DONE"
