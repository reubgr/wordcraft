#!/bin/bash
# pdf_opts6_restore.command — restore the shared Word profile's Track
# Changes Options to the pre-run snapshot after the v5 driver set the
# custom values but its restore call failed on the `line` identifier
# collision (pdf_opts5.applescript, -1728 in valOf). Uses pdf_opts6
# (fixed identifiers). Lock-held; verifies by full snapshot readback.
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_opts6_restore.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== OPTS6_RESTORE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-opts6-restore --wait 1200 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== OPTS6_RESTORE_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
echo "--- before"
runwd 90 osascript "$D/pdf_opts6.applescript" snapshot 2>&1
ORIG=$(printf 'SNAP\tinsColor=byAuthor\tinsMark=underline\tdelColor=byAuthor\tdelMark=strikethrough\tinsCellColor=lightBlue\tdelCellColor=pink\tmergedCellColor=lightYellow\tsplitCellColor=lightOrange\tmoveFromColor=green\tmoveFromMark=doubleStrikethrough\tmoveToColor=green\tmoveToMark=doubleUnderline\tcommentsColor=byAuthor\tprintComments=missing value')
echo "--- restore"
runwd 90 osascript "$D/pdf_opts6.applescript" restore "$ORIG" 2>&1
echo "RESTORE_RC=$?"
echo "=== OPTS6_RESTORE_DONE"
