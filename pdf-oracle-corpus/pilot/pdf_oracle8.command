#!/bin/bash
# pdf_oracle8.command — Word->PDF oracle PRODUCTION driver (#6153 full run).
# One group per invocation; the group is selected by the invoked name
# (pdf_oracle8_g1.command .. pdf_oracle8_g4.command). Each invocation is
# ONE WORD-LOCK claim (holder pdf-oracle-pilot) doing, for its entries:
# markup-default xN (profile options untouched, view forced to balloon
# first), then set-custom -> markup-custom xN -> restore + verify, then
# view restore, then clean xN. All v7 rules are mandatory and kept:
# ONE save per export, container path only (corpus out-dir reached by
# shell copy-out, never a Word save); name gate inside pdf_opts8 before
# every save; FINAL_DOCS must return to 0 after every export or
# production stops; options only via the application `settings`
# property (pdf_opts8.applescript = pdf_opts7 with the doExport
# AppleEvent timeout raised 120s -> 900s for the phase-4 stress docs).
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
SELF=$(basename "$0")
GRP="${SELF##*_g}"
GRP="${GRP%.command}"
OUT=$D/pdf_oracle8_g$GRP.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_ORACLE8 driver group $GRP at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
case "$GRP" in
  1) IDS="tc-p0-del-word tc-p0-identical tc-p0-ins-para tc-p0-ins-sentence tc-p0-ins-word tc-p0-sub-word tc-p1-case-only tc-p1-comment-added tc-p1-field-pair tc-p1-footnote-change tc-p1-formatting-only tc-p1-header-change tc-p1-move-para tc-p1-table-cell tc-p1-textbox-change tc-p1-whitespace-only tc-p2-base-del" ;;
  2) IDS="tc-p2-base-fmt tc-p2-base-ins tc-p2-both-del tc-p2-both-fmt tc-p2-both-ins tc-p2-both-samepara tc-p2-control tc-p2-rev-del tc-p2-rev-fmt tc-p2-rev-ins tc-p2m-cjk-change tc-p2m-combining-change tc-p2m-empty-both tc-p2m-empty-del tc-p2m-empty-ins tc-p2m-endnote-change" ;;
  3) IDS="tc-p2m-equation-change tc-p2m-image-replace tc-p2m-image-resize tc-p2m-list-convert tc-p2m-list-insert tc-p2m-list-restart tc-p2m-longpara tc-p2m-manyblocks tc-p2m-para-merge tc-p2m-para-split tc-p2m-rtl-change tc-p2m-section-header2 tc-p2m-section-margins tc-p2m-section-orient tc-p2m-style-para tc-p2m-style-run" ;;
  4) IDS="tc-p2m-table-colwidth tc-p2m-table-merge tc-p2m-table-nested tc-p2m-table-row-ins tc-p2m-table-rowswap tc-p3-churn-long tc-p4-budget-fewlong tc-p4-budget-manypairs tc-p4-dense-long tc-p4-dense-short tc-p4-ladder-10 tc-p4-ladder-100 tc-p4-ladder-1000 tc-p4-ladder-500 tc-p4-lcs-move tc-p4-lcs-move-ctl" ;;
  *) echo "UNKNOWN_GROUP $GRP — aborting"; echo "=== PDF_ORACLE8_DONE rc=group"; exit 2 ;;
esac
echo "GROUP_IDS: $IDS"
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
runwd 20 cp "$P/probe_src.txt" "$PROBE" && echo "CPROBE_WRITE_OK" || { echo "CPROBE_WRITE_FAIL — aborting"; echo "=== PDF_ORACLE8_DONE rc=probe"; exit 4; }
runwd 20 shasum -a 256 "$PROBE" >/dev/null && echo "CPROBE_READ_OK" || { echo "CPROBE_READ_FAIL — aborting"; echo "=== PDF_ORACLE8_DONE rc=probe"; exit 4; }
runwd 20 rm -f "$PROBE" && echo "CPROBE_RM_OK" || echo "CPROBE_RM_FAIL"
echo "--- pre-flight"
mdls -name kMDItemVersion "/Applications/Microsoft Word.app" 2>/dev/null
/usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/Microsoft Word.app/Contents/Info.plist" 2>/dev/null
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch "pdf-corpus-fullrun-g$GRP" --wait 7200 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  echo "=== PDF_ORACLE8_DONE rc=lock"
  exit 3
fi
echo "LOCK_ACQUIRED token=$TOK at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
echo "--- window census gate (no Grant File Access panels allowed)"
CENSUS=$(runwd 30 osascript "$D/dlg_prompt.applescript" census 2>&1)
echo "$CENSUS"
case "$CENSUS" in
  *"Grant File Access"*) echo "GATE_FAIL: grant panels present — aborting before Word work"; echo "=== PDF_ORACLE8_DONE rc=panels"; exit 6 ;;
esac
doc_names() { runwd 40 osascript "$D/pdf_docnames2.applescript" 2>/dev/null; }
BASE=$(doc_names)
echo "BASELINE $BASE"
case "$BASE" in
  "DOCS=0"*) : ;;
  *) echo "GATE_FAIL: baseline documents not zero — aborting"; echo "=== PDF_ORACLE8_DONE rc=baseline"; exit 6 ;;
esac
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "PING_OK"
end timeout' 2>&1
FIRST_ID=$(echo $IDS | cut -d" " -f1)
echo "FIRST_ID=$FIRST_ID"
echo "--- input hashes (source, before)"
for id in $IDS; do runwd 20 shasum -a 256 "$P/$id.docx" | sed "s|$P/||"; done
echo "--- staging (idempotent)"
for id in $IDS; do
  runwd 30 cp "$P/$id.docx" "$C/pdfpilot-$id.docx" && echo "STAGE_OK $id" || echo "STAGE_FAIL $id"
done
for id in $IDS; do runwd 20 shasum -a 256 "$C/pdfpilot-$id.docx" | sed "s|$C/||"; done
for id in $IDS; do rm -f "$C/pdfpilot-$id-word-clean.pdf" "$C/pdfpilot-$id-word-markup-default.pdf" "$C/pdfpilot-$id-word-markup-custom.pdf"; done
echo "STALE_CONTAINER_PDFS_REMOVED"
ORIG=$(runwd 90 osascript "$D/pdf_opts8.applescript" snapshot 2>&1)
echo "ORIGINAL $ORIG"
case "$ORIG" in
  *SNAPSHOT_FAIL*|"") echo "SNAPSHOT_FAILED — aborting"; echo "=== PDF_ORACLE8_DONE rc=snapshot"; exit 5 ;;
  *"insColor=ERR"*|*"delColor=ERR"*) echo "CUSTOM_BLOCKED: original ins/del colour unreadable — aborting"; echo "=== PDF_ORACLE8_DONE rc=snapshot"; exit 5 ;;
esac
echo "--- set view mode to balloon (profile default for markup-default)"
VS=$(runwd 90 osascript "$D/pdf_viewdoc.applescript" "$C/pdfpilot-$FIRST_ID.docx" balloon 2>&1)
echo "$VS"
case "$VS" in
  *"VIEW_SET_TO balloon"*) echo "VIEW_BALLOON_OK" ;;
  *) echo "VIEW_SET_FAILED — aborting (markup-default would be unfaithful)"; echo "=== PDF_ORACLE8_DONE rc=viewset"; exit 6 ;;
esac
STATE_OK=1
export_one() { # export_one MODE ID FNAME
  local mode=$1 id=$2 fname=$3 out rc fin
  mkdir -p "$P/out/$id"
  echo "=== EXPORT $mode $id at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  out=$(runwd 1000 osascript "$D/pdf_opts8.applescript" export "$mode" "$C/pdfpilot-$id.docx" "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" 2>&1)
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
echo "--- markup-default phase"
NEXP=0
for id in $IDS; do
  [ $STATE_OK -eq 1 ] && export_one markup-default "$id" word-markup-default.pdf
  NEXP=$((NEXP+1))
  if [ $NEXP -eq 1 ]; then
    FIRSTVB=$(grep -m1 'VIEW_BEFORE' "$OUT" | sed 's/.*VIEW_BEFORE //')
    echo "FIRST_DEFAULT_VIEW_BEFORE=$FIRSTVB"
    if [ "$FIRSTVB" != "balloon" ]; then echo "GATE_FAIL: markup-default view mode is $FIRSTVB, not balloon — stopping production"; STATE_OK=0; fi
  fi
done
ORIG_VIEW=$(grep -m1 'VIEW_BEFORE' "$OUT" | sed 's/.*VIEW_BEFORE //')
echo "ORIG_VIEW_MODE=$ORIG_VIEW"
CUSTOM_OK=0
if [ $STATE_OK -eq 1 ]; then
  echo "--- set-custom"
  SETOUT=$(runwd 90 osascript "$D/pdf_opts8.applescript" set-custom 2>&1)
  echo "$SETOUT"
  CUSTLINE=$(printf '%s\n' "$SETOUT" | grep '^CUSTOM' | head -1)
  if printf '%s' "$CUSTLINE" | grep -q "insColor=blue" && printf '%s' "$CUSTLINE" | grep -q "delColor=red" && printf '%s' "$CUSTLINE" | grep -q "insMark=underline" && printf '%s' "$CUSTLINE" | grep -q "delMark=strikethrough"; then
    CUSTOM_OK=1; echo "CUSTOM_SET_VERIFIED"
  else
    echo "CUSTOM_SET_INCOMPLETE — skipping custom phase, going to restore"
  fi
fi
if [ $CUSTOM_OK -eq 1 ]; then
  echo "--- markup-custom phase"
  for id in $IDS; do
    [ $STATE_OK -eq 1 ] && export_one markup-custom "$id" word-markup-custom.pdf
  done
fi
echo "--- restore options"
RESTORED=$(runwd 90 osascript "$D/pdf_opts8.applescript" restore "$ORIG" 2>&1)
echo "RESTORED_LINE $RESTORED"
O=$(printf '%s' "$ORIG" | cut -f2-)
R=$(printf '%s' "$RESTORED" | sed 's/^RESTORED//' | cut -f2-)
if [ "$O" = "$R" ]; then echo "RESTORE_MATCH"; else echo "RESTORE_MISMATCH orig=[$O] restored=[$R]"; fi
if [ -n "$ORIG_VIEW" ] && [ "$ORIG_VIEW" != "inline" ] && [ "$ORIG_VIEW" != "unknown" ]; then
  runwd 90 osascript "$D/pdf_viewdoc.applescript" "$C/pdfpilot-$FIRST_ID.docx" "$ORIG_VIEW" 2>&1
  echo "VIEW_RESTORE_RC=$?"
fi
if [ $STATE_OK -eq 1 ]; then
  echo "--- clean phase"
  for id in $IDS; do
    [ $STATE_OK -eq 1 ] && export_one clean "$id" word-clean.pdf
  done
fi
echo "--- final doc state"
doc_names
echo "--- input hashes (source, after)"
for id in $IDS; do runwd 20 shasum -a 256 "$P/$id.docx" | sed "s|$P/||"; done
echo "--- input hashes (staged copies, after)"
for id in $IDS; do runwd 20 shasum -a 256 "$C/pdfpilot-$id.docx" | sed "s|$C/||"; done
for id in $IDS; do rm -f "$C/pdfpilot-$id-word-clean.pdf" "$C/pdfpilot-$id-word-markup-default.pdf" "$C/pdfpilot-$id-word-markup-custom.pdf"; done
echo "CONTAINER_PDF_CLEAN_DONE"
echo "--- outputs"
for id in $IDS; do ls -la "$P/out/$id/" 2>/dev/null | grep -E "word-(clean|markup-default|markup-custom)\.pdf" | sed "s|^|$id/|"; done
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
echo "=== PDF_ORACLE8_DONE group $GRP state_ok=$STATE_OK custom_ok=$CUSTOM_OK at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
