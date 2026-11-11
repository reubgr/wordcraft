#!/bin/bash
# pdf_settingsprobe.command — probe 3 (#6153): the application class exposes
# the options object as property `settings` (wSet). Test get/set of Track
# Changes colours through `settings` (the pilot's earlier scripts used the
# bare class name `Word options`: sets -10006, gets -1728). Also test
# Word save-as PDF directly to the out-dir (non-container) path.
# Lock-held; sets are restored from the readback if any set succeeds;
# everything closed without saving; released clean.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_settingsprobe.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== SETTINGSPROBE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-settingsprobe --wait 900 2>/dev/null)
if [ -z "$TOK" ]; then echo "LOCK_NOT_ACQUIRED"; echo "=== SETTINGSPROBE_DONE rc=lock"; exit 3; fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
trap 'kill $HB 2>/dev/null; bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN release deferred"' EXIT
echo "--- formulations on settings object (no doc needed?)"
runwd 30 osascript <<'AS' 2>&1
with timeout of 25 seconds
	tell application "Microsoft Word"
		set o to settings
		set r to "GET_VIA_SETTINGS insColorEqBlue="
		if (inserted text color of o) = blue then set r to r & "true"
		if (inserted text color of o) = by author then set r to r & "byAuthor"
		return r
	end tell
end timeout
AS
runwd 30 osascript <<'AS' 2>&1
with timeout of 25 seconds
	tell application "Microsoft Word"
		set before to (inserted text color of settings) = by author
		set inserted text color of settings to blue
		set afterBlue to (inserted text color of settings) = blue
		set inserted text color of settings to by author
		set restored to (inserted text color of settings) = by author
		return "SET_VIA_SETTINGS beforeByAuthor=" & before & " afterBlue=" & afterBlue & " restoredByAuthor=" & restored
	end tell
end timeout
AS
echo "--- open staged copy + save-as PDF to out-dir and container"
mkdir -p "$P/out/tc-p2m-multiauthor"
runwd 90 osascript <<'AS' 2>&1
with timeout of 80 seconds
	tell application "Microsoft Word"
		open "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/pdfpilot-tc-p2m-multiauthor.docx"
		set doc to active document
		set show revisions of doc to true
		set print revisions of doc to true
		save as doc file name "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/probe-settings.pdf" file format format PDF
		save as doc file name "/Users/muse/word-oracle/pdf-pilot/out/tc-p2m-multiauthor/probe-settings.pdf" file format format PDF
		close doc saving no
		return "SAVE_TEST_DONE docs=" & (count of documents)
	end tell
end timeout
AS
ls -la "$P/out/tc-p2m-multiauthor/" 2>&1
echo "=== SETTINGSPROBE_DONE"
