#!/bin/bash
# pdf_recover.command — supervised Word recovery (#6153): the restarted
# session's Word (pid seen at 17:36, parented to launchd) answers count
# queries but silently no-ops dictionary opens, has no windows, and shell
# access to its container blocks (all lanes affected — vba_matrix died
# on a container write too). With docs=0, quit + relaunch Word inside
# the GUI session under the WORD-LOCK, then retest the three primitives:
# container shell read, dictionary open of a staged copy, options snapshot.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
OUT=$D/pdf_recover.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== RECOVER at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-recover --wait 900 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== RECOVER_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
cnt() { runwd 12 osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return (count of documents)
end timeout' 2>&1; }
echo "docs_before=$(cnt)"
WPID=$(pgrep -f "Microsoft Word.app/Contents/MacOS/Microsoft Word" | head -1)
echo "word_pid=$WPID"
runwd 30 osascript -e 'with timeout of 20 seconds
tell application "Microsoft Word" to quit
end timeout' 2>&1
echo "quit_rc=$?"
n=0
while [ $n -lt 15 ]; do
  sleep 2; n=$((n+1))
  pgrep -f "Microsoft Word.app/Contents/MacOS/Microsoft Word" >/dev/null || break
done
if pgrep -f "Microsoft Word.app/Contents/MacOS/Microsoft Word" >/dev/null; then
  echo "WORD_STILL_RUNNING after quit; sending TERM"
  pkill -TERM -f "Microsoft Word.app/Contents/MacOS/Microsoft Word"
  sleep 5
fi
pgrep -f "Microsoft Word.app/Contents/MacOS/Microsoft Word" >/dev/null && echo "WORD_PROCESS_PRESENT" || echo "WORD_PROCESS_GONE"
echo "--- relaunch via LaunchServices"
open -a "Microsoft Word"
n=0
while [ $n -lt 20 ]; do
  sleep 3; n=$((n+1))
  R=$(cnt)
  echo "ping$n -> $R"
  [ "$R" = "0" ] && break
done
runwd 15 osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (name of every window)' 2>&1
NEWPID=$(pgrep -f "Microsoft Word.app/Contents/MacOS/Microsoft Word" | head -1)
echo "word_pid_after=$NEWPID"
echo "--- retest 1: container shell read (watchdog 15s)"
runwd 15 shasum -a 256 "$C/pdfpilot-tc-p2m-multiauthor.docx" 2>&1 | sed "s|$C/||"
echo "--- retest 2: dictionary open of staged copy"
runwd 60 osascript -e 'with timeout of 50 seconds
tell application "Microsoft Word"
open "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/pdfpilot-tc-p2m-multiauthor.docx"
return "OPENRET docs=" & (count of documents) & " active=" & (name of active document)
end tell
end timeout' 2>&1
echo "--- retest 3: options snapshot (doc state as-is)"
runwd 30 osascript "$D/pdf_opts4.applescript" snapshot 2>&1
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word"
repeat while (count of documents) > 0
close document 1 saving no
end repeat
return "FINAL docs=" & (count of documents)
end tell
end timeout' 2>&1
echo "=== RECOVER_DONE"
