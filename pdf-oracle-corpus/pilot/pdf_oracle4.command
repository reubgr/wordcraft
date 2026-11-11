#!/bin/bash
# pdf_oracle4.command — Word->PDF oracle pilot RESUMPTION driver v4 (#6153).
# v3 died at staging: in the restarted GUI session a shell `cp` INTO Word's
# container blocks indefinitely (probe-confirmed twice, no TCC prompt
# visible). v4 therefore never writes into the container from the shell:
# documents are opened via LaunchServices from ~/word-oracle/pdf-pilot/
# (grant-free, the oracle-open-batch pattern), Word itself saves the PDFs
# (container path first — Word-internal, grant-free — then the out-dir
# path as a second save), and any container copy-out is watchdog-bounded.
# Every osascript call is watchdog-bounded (hang rule: < 180 s).
# Options are snapshotted equality-based (pdf_opts4.applescript), custom
# values are set + read back, originals are restored and re-read.
# Lock lifecycle identical to v2/v3: word_lock.sh only, clean release.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_oracle4.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_ORACLE4 driver at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
caffeinate -d -i -m -s -u &
CAF=$!
runwd() { # runwd SECONDS cmd... -> rc 124 on watchdog kill
  local limit=$1; shift
  "$@" &
  local pid=$! n=0
  while kill -0 $pid 2>/dev/null; do
    if [ $n -ge $limit ]; then kill -9 $pid 2>/dev/null; wait $pid 2>/dev/null; echo "WATCHDOG_KILL after ${limit}s: $*" ; return 124; fi
    sleep 1; n=$((n+1))
  done
  wait $pid; return $?
}
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-corpus-pilot-resume4 --wait 1500 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  kill $CAF 2>/dev/null
  echo "=== PDF_ORACLE4_DONE rc=lock"
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
T0=$(python3 -c 'import time; print(time.time())')
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "PING docs=" & (count of documents)
end timeout' 2>&1
T1=$(python3 -c 'import time; print(time.time())')
python3 -c "print(f'DICT_PING_WALL {$T1-$T0:.2f}')"
runwd 15 osascript -e 'tell application "System Events" to tell process "Microsoft Word" to return "WINDOWS " & (count of windows)' 2>&1
IDS="tc-p0-del-para tc-p2m-multiauthor tc-p2m-table-row-del"
echo "--- input hashes (source, before)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
doc_count() { runwd 12 osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return (count of documents)
end timeout' 2>/dev/null; }
open_doc() { # open_doc ID -> 0 when the doc is open and active name matches
  local id=$1 n=0 cnt name
  open "$P/$id.docx"
  while [ $n -lt 30 ]; do
    sleep 2; n=$((n+1))
    cnt=$(doc_count)
    if [ "$cnt" = "1" ]; then
      name=$(runwd 12 osascript -e 'with timeout of 8 seconds
tell application "Microsoft Word" to return (name of active document)
end timeout' 2>/dev/null)
      echo "OPENED $id name=[$name]"
      case "$name" in *"$id"*) return 0;; *) echo "OPEN_NAME_MISMATCH $id"; return 1;; esac
    fi
  done
  echo "OPEN_TIMEOUT $id count=$cnt"
  return 1
}
close_doc() { # close active document without saving; verify count 0
  runwd 30 osascript -e 'with timeout of 20 seconds
tell application "Microsoft Word" to close active document saving no
end timeout' 2>&1
  local n=0 cnt
  while [ $n -lt 10 ]; do
    cnt=$(doc_count)
    [ "$cnt" = "0" ] && { echo "CLOSED_OK"; return 0; }
    sleep 2; n=$((n+1))
  done
  echo "CLOSE_FAIL count=$cnt"
  return 1
}
export_one() { # export_one MODE ID
  local mode=$1 id=$2 fname
  case $mode in
    markup-custom) fname=word-markup-custom.pdf ;;
    clean)         fname=word-clean.pdf ;;
  esac
  mkdir -p "$P/out/$id"
  if ! open_doc "$id"; then echo "EXPORT_SKIP $mode $id (open failed)"; return 1; fi
  runwd 120 osascript "$D/pdf_opts4.applescript" export "$mode" "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" 2>&1
  echo "EXPORT_RC=$? $mode $id"
  close_doc
  # copy-out fallback if Word's direct out-dir save did not land the file
  # (no stat on the container path — watchdog-bounded cp only)
  if [ ! -f "$P/out/$id/$fname" ]; then
    runwd 20 cp "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" && echo "COPYOUT_OK $id $fname" || echo "COPYOUT_FAIL $id $fname"
  fi
  ls -la "$P/out/$id/$fname" 2>/dev/null || echo "RESULT_MISSING $id $fname"
}
# ---- options: snapshot original ----
ORIG=$(runwd 60 osascript "$D/pdf_opts4.applescript" snapshot 2>&1)
echo "ORIGINAL $ORIG"
# ---- custom sets + readback ----
runwd 60 osascript "$D/pdf_opts4.applescript" set-custom 2>&1
# ---- markup-custom x3 ----
ORIG_VIEW=""
for id in $IDS; do export_one markup-custom "$id"; done
# capture the pre-custom view mode from the first export's VIEW_BEFORE is in
# the log above; the driver extracts it for the view restore below.
ORIG_VIEW=$(grep -m1 "VIEW_BEFORE" "$OUT" | sed 's/.*VIEW_BEFORE //')
echo "ORIG_VIEW_MODE=$ORIG_VIEW"
# ---- restore options + readback ----
if [ -n "$ORIG" ]; then
  echo "RESTORED $(runwd 60 osascript "$D/pdf_opts4.applescript" restore "$ORIG" 2>&1)"
else
  echo "RESTORE_SKIPPED (no ORIGINAL snapshot)"
fi
# ---- view-mode restore (open first doc, set view back, close) ----
if [ -n "$ORIG_VIEW" ] && [ "$ORIG_VIEW" != "inline" ] && [ "$ORIG_VIEW" != "unknown" ]; then
  if open_doc tc-p0-del-para; then
    runwd 30 osascript "$D/pdf_opts4.applescript" set-view "$ORIG_VIEW" 2>&1
    close_doc
  fi
fi
# ---- clean x3 ----
for id in $IDS; do export_one clean "$id"; done
echo "--- input hashes (source, after)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "FINAL docs=" & (count of documents)
end timeout' 2>&1
kill $CAF 2>/dev/null
echo "=== PDF_ORACLE4_DONE"
