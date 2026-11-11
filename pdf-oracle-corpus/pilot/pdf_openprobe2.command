#!/bin/bash
# pdf_openprobe2.command — supervised probe 2 (#6153): dictionary opens.
# (a) hash the run-3 staged container copies (watchdog-bounded shell read);
# (b) dictionary-open a staged copy; (c) options snapshot WITH a doc open;
# (d) dictionary-open an original from ~/word-oracle/pdf-pilot/ directly.
# Lock-held throughout; everything closed without saving; released clean.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_openprobe2.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== OPENPROBE2 at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-openprobe2 --wait 900 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== OPENPROBE2_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
cnt() { runwd 12 osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return (count of documents)
end timeout' 2>&1; }
echo "--- staged container hashes (expect 5c7062.. / 5f5667.. / 50a0d2..)"
for id in tc-p0-del-para tc-p2m-multiauthor tc-p2m-table-row-del; do
  runwd 15 shasum -a 256 "$C/pdfpilot-$id.docx" 2>&1 | sed "s|$C/||"
done
echo "--- dictionary open of staged copy (tc-p2m-multiauthor)"
runwd 60 osascript -e 'with timeout of 50 seconds
tell application "Microsoft Word"
open "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/pdfpilot-tc-p2m-multiauthor.docx"
return "OPENRET docs=" & (count of documents) & " active=" & (name of active document)
end tell
end timeout' 2>&1
echo "count=$(cnt)"
echo "--- snapshot WITH doc open (if any):"
runwd 30 osascript "$D/pdf_opts4.applescript" snapshot 2>&1
echo "--- set-custom attempt WITH doc open:"
runwd 30 osascript "$D/pdf_opts4.applescript" set-custom 2>&1
echo "--- restore options from a by-author probe line (no-op if sets failed):"
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to close active document saving no
end timeout' 2>&1
sleep 2; echo "count_after_close=$(cnt)"
echo "--- dictionary open of ORIGINAL from pdf-pilot (ungranted path test)"
runwd 60 osascript -e 'with timeout of 50 seconds
tell application "Microsoft Word"
open "/Users/muse/word-oracle/pdf-pilot/tc-p0-del-para.docx"
return "OPENRET docs=" & (count of documents) & " active=" & (name of active document)
end tell
end timeout' 2>&1
echo "count=$(cnt)"
runwd 15 osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (name of every window)' 2>&1
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word"
repeat while (count of documents) > 0
close document 1 saving no
end repeat
return "FINAL docs=" & (count of documents)
end tell
end timeout' 2>&1
echo "=== OPENPROBE2_DONE"
