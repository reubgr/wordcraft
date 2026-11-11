#!/bin/bash
# pdf_oracle7.command — Word->PDF oracle pilot driver v7 (#6153).
# Scope: the 5 exports the poisoned v6 run did not validly produce —
# markup-custom for tc-p2m-multiauthor + tc-p2m-table-row-del, and
# clean for all three pilot entries. (tc-p0-del-para markup-custom was
# validly produced by v6 export 1 and is kept.) v7 rules, learned from
# v6: ONE save per export, container path only (a save to the
# un-granted out dir raises a Grant File Access panel whose modal stack
# wedges opens/closes); the export script name-gates the active
# document before saving; this driver additionally gates on baseline
# docs = 0, a grant-panel-free window census, and FINAL_DOCS returning
# to baseline after every export — any gate failure stops production
# and goes to restore/release. Options only via `settings`
# (pdf_opts7.applescript); originals snapshotted first, custom values
# verified by readback, originals restored and re-read at the end.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_oracle7.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_ORACLE7 driver at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
caffeinate -d -i -m -s -u &
CAF=$!
runwd() {
  local limit=$1; shift
  "$@" &
  local pid=$! n=0
  while kill -0 $pid 2>/dev/null; do
    if [ $n -ge $limit ]; then kill -9 $pid 2>/dev/null; wait $pid 2>/dev/null; echo "WATCHDOG_KILL after ${limit}s: $*" ; return 124; fi
    sleep 1; n=$((n+1))
  done
  wait $pid; return $?
}
TOK=""
RELEASED=0
trap 'kill $CAF 2>/dev/null; kill $HB 2>/dev/null; if [ $RELEASED -eq 0 ] && [ -n "$TOK" ]; then bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1 || echo "TRAP_RELEASE_DEFERRED"; fi' EXIT
echo "--- container probe (watchdog 20s)"
PROBE="$C/.pdfpilot_probe_$$"
printf 'pdf-oracle-probe %s\n' "$(date -u +%s)" > "$P/probe_src.txt"
runwd 20 cp "$P/probe_src.txt" "$PROBE" && echo "CPROBE_WRITE_OK" || { echo "CPROBE_WRITE_FAIL — aborting"; echo "=== PDF_ORACLE7_DONE rc=probe"; exit 4; }
runwd 20 shasum -a 256 "$PROBE" >/dev/null && echo "CPROBE_READ_OK" || { echo "CPROBE_READ_FAIL — aborting"; echo "=== PDF_ORACLE7_DONE rc=probe"; exit 4; }
runwd 20 rm -f "$PROBE" && echo "CPROBE_RM_OK" || echo "CPROBE_RM_FAIL"
echo "--- pre-flight"
mdls -name kMDItemVersion "/Applications/Microsoft Word.app" 2>/dev/null
/usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/Microsoft Word.app/Contents/Info.plist" 2>/dev/null
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-corpus-pilot-resume7 --wait 3600 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  echo "=== PDF_ORACLE7_DONE rc=lock"
  exit 3
fi
echo "LOCK_ACQUIRED token=$TOK at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
echo "--- window census gate (no Grant File Access panels allowed)"
CENSUS=$(runwd 30 osascript "$D/dlg_prompt.applescript" census 2>&1)
echo "$CENSUS"
case "$CENSUS" in
  *"Grant File Access"*) echo "GATE_FAIL: grant panels present — aborting before Word work"; echo "=== PDF_ORACLE7_DONE rc=panels"; exit 6 ;;
esac
doc_names() { runwd 40 osascript "$D/pdf_docnames2.applescript" 2>/dev/null; }
BASE=$(doc_names)
echo "BASELINE $BASE"
case "$BASE" in
  "DOCS=0"*) : ;;
  *) echo "GATE_FAIL: baseline documents not zero — aborting"; echo "=== PDF_ORACLE7_DONE rc=baseline"; exit 6 ;;
esac
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "PING_OK"
end timeout' 2>&1
IDS_ALL="tc-p0-del-para tc-p2m-multiauthor tc-p2m-table-row-del"
echo "--- input hashes (source, before)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
echo "--- staging (idempotent)"
for id in $IDS_ALL; do
  runwd 30 cp "$P/$id.docx" "$C/pdfpilot-$id.docx" && echo "STAGE_OK $id" || echo "STAGE_FAIL $id"
done
runwd 30 shasum -a 256 "$C"/pdfpilot-tc-*.docx | sed "s|$C/||"
runwd 20 rm -f "$C"/pdfpilot-tc-*-word-*.pdf && echo "STALE_CONTAINER_PDFS_REMOVED"
ORIG=$(runwd 90 osascript "$D/pdf_opts7.applescript" snapshot 2>&1)
echo "ORIGINAL $ORIG"
case "$ORIG" in
  *SNAPSHOT_FAIL*|"") echo "SNAPSHOT_FAILED — aborting"; echo "=== PDF_ORACLE7_DONE rc=snapshot"; exit 5 ;;
  *"insColor=ERR"*|*"delColor=ERR"*) echo "CUSTOM_BLOCKED: original ins/del colour unreadable — aborting (clean-only rerun needs a new driver decision)"; echo "=== PDF_ORACLE7_DONE rc=snapshot"; exit 5 ;;
esac
echo "--- set-custom"
SETOUT=$(runwd 90 osascript "$D/pdf_opts7.applescript" set-custom 2>&1)
echo "$SETOUT"
CUSTLINE=$(printf '%s\n' "$SETOUT" | grep '^CUSTOM' | head -1)
CUSTOM_OK=0
if printf '%s' "$CUSTLINE" | grep -q "insColor=blue" && printf '%s' "$CUSTLINE" | grep -q "delColor=red" && printf '%s' "$CUSTLINE" | grep -q "insMark=underline" && printf '%s' "$CUSTLINE" | grep -q "delMark=strikethrough"; then
  CUSTOM_OK=1; echo "CUSTOM_SET_VERIFIED"
else
  echo "CUSTOM_SET_INCOMPLETE — aborting to restore"
fi
STATE_OK=1
export_one() { # export_one MODE ID FNAME
  local mode=$1 id=$2 fname=$3 out rc fin
  mkdir -p "$P/out/$id"
  out=$(runwd 150 osascript "$D/pdf_opts7.applescript" export "$mode" "$C/pdfpilot-$id.docx" "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" 2>&1)
  rc=$?
  echo "$out"
  echo "EXPORT_RC=$rc $mode $id"
  case "$out" in
    *OPEN_NAME_MISMATCH*|*OPEN_FAIL*|*EXPORT_SETUP_FAIL*|*SAVE_CONTAINER_FAIL*) echo "EXPORT_INVALID $mode $id"; STATE_OK=0; return 1 ;;
  esac
  fin=$(printf '%s' "$out" | grep -o 'FINAL_DOCS [0-9]*' | tail -1)
  if [ "$fin" != "FINAL_DOCS 0" ]; then echo "EXPORT_STATE_BAD $mode $id ($fin)"; STATE_OK=0; return 1; fi
  runwd 20 cp "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" && echo "COPYOUT_OK $id $fname" || echo "COPYOUT_FAIL $id $fname"
  ls -la "$P/out/$id/$fname" 2>/dev/null || echo "RESULT_MISSING $id $fname"
  runwd 15 shasum -a 256 "$P/out/$id/$fname" 2>/dev/null | sed "s|$P/out/||"
}
if [ $CUSTOM_OK -eq 1 ]; then
  for id in tc-p2m-multiauthor tc-p2m-table-row-del; do
    [ $STATE_OK -eq 1 ] && export_one markup-custom "$id" word-markup-custom.pdf
  done
  ORIG_VIEW=$(grep -m1 "VIEW_BEFORE" "$OUT" | sed 's/.*VIEW_BEFORE //')
  echo "ORIG_VIEW_MODE=$ORIG_VIEW"
fi
echo "--- restore options"
RESTORED=$(runwd 90 osascript "$D/pdf_opts7.applescript" restore "$ORIG" 2>&1)
echo "RESTORED_LINE $RESTORED"
O=$(printf '%s' "$ORIG" | cut -f2-)
R=$(printf '%s' "$RESTORED" | sed 's/^RESTORED//' | cut -f2-)
if [ "$O" = "$R" ]; then echo "RESTORE_MATCH"; else echo "RESTORE_MISMATCH orig=[$O] restored=[$R]"; fi
if [ -n "$ORIG_VIEW" ] && [ "$ORIG_VIEW" != "inline" ] && [ "$ORIG_VIEW" != "unknown" ]; then
  runwd 90 osascript "$D/pdf_viewdoc.applescript" "$C/pdfpilot-tc-p0-del-para.docx" "$ORIG_VIEW" 2>&1
  echo "VIEW_RESTORE_RC=$?"
fi
echo "--- clean x3"
for id in $IDS_ALL; do
  [ $STATE_OK -eq 1 ] && export_one clean "$id" word-clean.pdf
done
echo "--- final doc state"
doc_names
echo "--- input hashes (source, after)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
echo "--- input hashes (staged copies, after)"
runwd 30 shasum -a 256 "$C"/pdfpilot-tc-*.docx | sed "s|$C/||"
runwd 20 rm -f "$C"/pdfpilot-tc-*-word-*.pdf && echo "CONTAINER_PDF_CLEAN_OK" || echo "CONTAINER_PDF_CLEAN_FAIL"
echo "--- outputs"
ls -la "$P"/out/*/word-clean.pdf "$P"/out/*/word-markup-custom.pdf 2>&1
kill $HB 2>/dev/null
if bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1; then
  RELEASED=1; echo "LOCK_RELEASED plain"
else
  NAMES=$(doc_names)
  echo "PLAIN_RELEASE_REFUSED docs: $NAMES"
  case "$NAMES" in
    *pdfpilot*) echo "RELEASE_DEFERRED: pilot documents still open" ;;
    *DOCS=*) bash "$D/word_lock.sh" release --force --token "$TOK" >/dev/null 2>&1 && { RELEASED=1; echo "LOCK_RELEASED force (no pilot documents remain)"; } || echo "RELEASE_FAILED" ;;
    *) echo "RELEASE_DEFERRED: census unreadable" ;;
  esac
fi
kill $CAF 2>/dev/null
echo "=== PDF_ORACLE7_DONE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
