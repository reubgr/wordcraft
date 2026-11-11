#!/bin/bash
# pdf_oracle.command — Word->PDF oracle pilot driver (storytold/wordcraft#6153).
# Runs under the resident queue-runner (Terminal session grants).
# Acquires the WORD-LOCK itself via word_lock.sh (bounded wait), stages
# inputs into Word's container, runs pdf_oracle.applescript, copies PDFs
# out to ~/word-oracle/pdf-pilot/out/, verifies input hashes unchanged,
# and releases the lock only when Word is back to the clean state.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_oracle.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_ORACLE driver at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
caffeinate -d -i -m -s -u &
CAF=$!
# ---- WORD-LOCK ----
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-corpus-pilot --wait 1500 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  kill $CAF 2>/dev/null
  echo "=== PDF_ORACLE_DONE rc=lock"
  exit 3
fi
echo "LOCK_ACQUIRED token=$TOK"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
wl_release() {
  kill $HB 2>/dev/null
  bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "WARN: lock release deferred (Word state not clean)"
}
trap 'kill $CAF 2>/dev/null; wl_release' EXIT
# ---- pre-flight gate ----
echo "--- pre-flight"
mdls -name kMDItemVersion "/Applications/Microsoft Word.app" 2>/dev/null
/usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/Microsoft Word.app/Contents/Info.plist" 2>/dev/null
pmset -g assertions 2>/dev/null | grep -i "PreventUserIdleDisplaySleep" | head -2
osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "PING docs=" & (count of documents)
end timeout' 2>&1
osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (count of windows)' 2>&1
# ---- stage inputs ----
mkdir -p "$P/out"
IDS="tc-p0-del-para tc-p2m-multiauthor tc-p2m-table-row-del"
for id in $IDS; do
  cp "$P/$id.docx" "$C/pdfpilot-$id.docx" || echo "STAGE_FAIL $id"
done
echo "--- staged hashes (before)"
shasum -a 256 "$C"/pdfpilot-*.docx | sed "s|$C/||"
# ---- build job (order: all markup-default, all markup-custom, all clean) ----
: > "$C/pdf_job.tsv"
for mode in markup-default markup-custom clean; do
  for id in $IDS; do
    case $mode in
      markup-default) f=word-markup-default.pdf ;;
      markup-custom)  f=word-markup-custom.pdf ;;
      clean)          f=word-clean.pdf ;;
    esac
    printf '%s\t%s\t%s\n' "$mode" "$C/pdfpilot-$id.docx" "$C/pdfpilot-$id-$f" >> "$C/pdf_job.tsv"
  done
done
cat "$C/pdf_job.tsv"
rm -f "$C/pdf_log.txt"
T0=$(python3 -c 'import time; print(time.time())')
osascript "$D/pdf_oracle.applescript" 2>&1
T1=$(python3 -c 'import time; print(time.time())')
python3 -c "print(f'APPLESCRIPT_WALL {$T1-$T0:.2f}')"
echo "--- macro log:"
cat "$C/pdf_log.txt" 2>/dev/null
# ---- copy results out ----
for id in $IDS; do
  mkdir -p "$P/out/$id"
  for f in word-clean.pdf word-markup-default.pdf word-markup-custom.pdf; do
    if [ -f "$C/pdfpilot-$id-$f" ]; then
      cp "$C/pdfpilot-$id-$f" "$P/out/$id/$f" && echo "RESULT_OK $id $f"
    else
      echo "RESULT_MISSING $id $f"
    fi
  done
done
echo "--- staged hashes (after)"
shasum -a 256 "$C"/pdfpilot-*.docx | sed "s|$C/||"
# ---- clean-state check ----
osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "FINAL docs=" & (count of documents)
end timeout' 2>&1
kill $CAF 2>/dev/null
echo "=== PDF_ORACLE_DONE"
