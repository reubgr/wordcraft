#!/bin/bash
# pdf_openprobe.command — supervised probe (#6153): which LaunchServices
# open variant opens a pilot docx in Word, and does the Word options
# snapshot succeed WITH a document open (v4 got -1728 with none open)?
# Acquires the WORD-LOCK itself; closes everything it opens; releases.
D=/Users/muse/compare-vs-word/dialog
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_openprobe.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== OPENPROBE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-openprobe --wait 900 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== OPENPROBE_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
cnt() { osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return (count of documents)
end timeout' 2>&1; }
try_open() { # try_open LABEL cmd...
  local label=$1; shift
  echo "--- TRY $label : $*"
  "$@" "$P/tc-p0-del-para.docx" 2>&1
  echo "open_rc=$?"
  local n=0 c
  while [ $n -lt 10 ]; do sleep 2; n=$((n+1)); c=$(cnt); echo "poll$n count=$c"; [ "$c" = "1" ] && break; done
  if [ "$c" = "1" ]; then
    osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return "ACTIVE " & (name of active document)
end timeout' 2>&1
    osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (name of every window)' 2>&1
    echo "--- snapshot WITH doc open:"
    osascript "$D/pdf_opts4.applescript" snapshot 2>&1
    osascript -e 'with timeout of 15 seconds
tell application "Microsoft Word" to close active document saving no
end timeout' 2>&1
    sleep 2; echo "after_close count=$(cnt)"
    return 0
  fi
  osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (name of every window)' 2>&1
  return 1
}
try_open PLAIN open || try_open APP open -a "Microsoft Word"
echo "=== OPENPROBE_DONE"
