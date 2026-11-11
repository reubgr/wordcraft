#!/bin/bash
# pdf_oracle5.command — Word->PDF oracle pilot RESUMPTION driver v5 (#6153).
# Resumes the frozen pilot after the container-write fault cleared by
# observation (watchdog probe passed 2026-11-10 17:23 CST). Route = the
# run-3 production route, unchanged: stage inputs into Word's container as
# pdfpilot-<id>.docx (shell cp, watchdog-bounded), dictionary open of the
# staged copy, save as format PDF to the container, close saving no,
# shell copy-out. Options work goes ONLY through the application
# `settings` property (pdf_opts5.applescript; probe-3b-proven) — never
# the bare `Word options` reference. Custom vector: insertions blue /
# underline, deletions red / strikethrough, balloons inline (view
# revisions mode = in line), then every changed value restored from the
# pre-run snapshot and verified by readback. Clean vector: revisions
# view final, show/print revisions false. Every external call is
# watchdog-bounded (3-minute hang rule); WORD-LOCK via word_lock.sh
# only, holder pdf-oracle-pilot, heartbeat 45 s.
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_oracle5.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== PDF_ORACLE5 driver at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=""
RELEASED=0
wl_release_plain() {
  [ -z "$TOK" ] && return 0
  bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1
}
trap 'kill $CAF 2>/dev/null; kill $HB 2>/dev/null; if [ $RELEASED -eq 0 ]; then wl_release_plain || echo "TRAP_RELEASE_DEFERRED"; fi' EXIT
# ---- container health probe BEFORE any lock or Word touch ----
echo "--- container probe (watchdog 20s)"
PROBE="$C/.pdfpilot_probe_$$"
printf 'pdf-oracle-probe %s\n' "$(date -u +%s)" > "$P/probe_src.txt"
runwd 20 cp "$P/probe_src.txt" "$PROBE" && echo "CPROBE_WRITE_OK" || { echo "CPROBE_WRITE_FAIL — aborting before lock"; echo "=== PDF_ORACLE5_DONE rc=probe"; exit 4; }
runwd 20 shasum -a 256 "$PROBE" && echo "CPROBE_READ_OK" || { echo "CPROBE_READ_FAIL — aborting before lock"; echo "=== PDF_ORACLE5_DONE rc=probe"; exit 4; }
runwd 20 rm -f "$PROBE" && echo "CPROBE_RM_OK" || echo "CPROBE_RM_FAIL (probe file left in container)"
# ---- pre-flight ----
echo "--- pre-flight"
mdls -name kMDItemVersion "/Applications/Microsoft Word.app" 2>/dev/null
/usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/Microsoft Word.app/Contents/Info.plist" 2>/dev/null
pmset -g assertions 2>/dev/null | grep -i "PreventUserIdleDisplaySleep" | head -2
ls -la /dev/console
# ---- acquire ----
TOK=$(bash "$D/word_lock.sh" acquire --holder pdf-oracle-pilot --batch pdf-corpus-pilot-resume5 --wait 3600 2>/dev/null)
RC=$?
if [ $RC -ne 0 ] || [ -z "$TOK" ]; then
  echo "LOCK_NOT_ACQUIRED rc=$RC — Word untouched"
  echo "=== PDF_ORACLE5_DONE rc=lock"
  exit 3
fi
echo "LOCK_ACQUIRED token=$TOK at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
doc_names() { runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word"
set out to "DOCS=" & (count of documents)
repeat with d in documents
set out to out & " | " & (name of d)
end repeat
return out
end timeout' 2>/dev/null; }
echo "--- baseline"
doc_names
PING0=$(python3 -c 'import time; print(time.time())')
runwd 15 osascript -e 'with timeout of 10 seconds
tell application "Microsoft Word" to return "PING_OK"
end timeout' 2>&1
PING1=$(python3 -c 'import time; print(time.time())')
python3 -c "print(f'DICT_PING_WALL {$PING1-$PING0:.2f}')"
IDS="tc-p0-del-para tc-p2m-multiauthor tc-p2m-table-row-del"
echo "--- input hashes (source, before)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
# ---- stage inputs into the container (run-3 route) ----
echo "--- staging"
for id in $IDS; do
  runwd 30 cp "$P/$id.docx" "$C/pdfpilot-$id.docx" && echo "STAGE_OK $id" || echo "STAGE_FAIL $id"
done
runwd 30 shasum -a 256 "$C"/pdfpilot-tc-*.docx | sed "s|$C/||"
# ---- exfil proof: copy one staged file back OUT of the container ----
runwd 20 cp "$C/pdfpilot-tc-p0-del-para.docx" "$P/exfil-proof.docx" && echo "EXFIL_PROOF_OK" || echo "EXFIL_PROOF_FAIL"
runwd 15 shasum -a 256 "$P/exfil-proof.docx" 2>/dev/null | sed "s|$P/||"
rm -f "$P/exfil-proof.docx"
# ---- options snapshot (original) ----
ORIG=$(runwd 90 osascript "$D/pdf_opts5.applescript" snapshot 2>&1)
echo "ORIGINAL $ORIG"
case "$ORIG" in
  *SNAPSHOT_FAIL*|"") echo "SNAPSHOT_FAILED — no sets, no exports; aborting clean"; echo "=== PDF_ORACLE5_DONE rc=snapshot"; exit 5 ;;
esac
ORIG_VIEW=""
CUSTOM_OK=0
case "$ORIG" in
  *"insColor=ERR"*|*"delColor=ERR"*) echo "CUSTOM_SKIPPED: original ins/del colour unreadable, cannot guarantee restore" ;;
  *)
    echo "--- set-custom"
    SETOUT=$(runwd 90 osascript "$D/pdf_opts5.applescript" set-custom 2>&1)
    echo "$SETOUT"
    CUSTLINE=$(printf '%s\n' "$SETOUT" | grep '^CUSTOM' | head -1)
    if printf '%s' "$CUSTLINE" | grep -q "insColor=blue" && printf '%s' "$CUSTLINE" | grep -q "delColor=red" && printf '%s' "$CUSTLINE" | grep -q "insMark=underline" && printf '%s' "$CUSTLINE" | grep -q "delMark=strikethrough"; then
      CUSTOM_OK=1; echo "CUSTOM_SET_VERIFIED"
    else
      echo "CUSTOM_SET_INCOMPLETE — custom exports skipped"
    fi
    ;;
esac
export_one() { # export_one MODE ID FNAME
  local mode=$1 id=$2 fname=$3
  mkdir -p "$P/out/$id"
  runwd 170 osascript "$D/pdf_opts5.applescript" export "$mode" "$C/pdfpilot-$id.docx" "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" 2>&1
  echo "EXPORT_RC=$? $mode $id"
  if [ ! -f "$P/out/$id/$fname" ]; then
    runwd 20 cp "$C/pdfpilot-$id-$fname" "$P/out/$id/$fname" && echo "COPYOUT_OK $id $fname" || echo "COPYOUT_FAIL $id $fname"
  else
    echo "OUT_SAVE_LANDED $id $fname"
  fi
  ls -la "$P/out/$id/$fname" 2>/dev/null || echo "RESULT_MISSING $id $fname"
  runwd 15 shasum -a 256 "$P/out/$id/$fname" 2>/dev/null | sed "s|$P/out/||"
}
# ---- markup-custom x3 ----
if [ $CUSTOM_OK -eq 1 ]; then
  for id in $IDS; do export_one markup-custom "$id" word-markup-custom.pdf; done
  ORIG_VIEW=$(grep -m1 "VIEW_BEFORE" "$OUT" | sed 's/.*VIEW_BEFORE //')
  echo "ORIG_VIEW_MODE=$ORIG_VIEW"
fi
# ---- restore options + readback (always, if any set was attempted) ----
if [ $CUSTOM_OK -eq 1 ] || printf '%s' "$SETOUT" | grep -q "SET_OK"; then
  RESTORED=$(runwd 90 osascript "$D/pdf_opts5.applescript" restore "$ORIG" 2>&1)
  echo "RESTORED_LINE $RESTORED"
  O=$(printf '%s' "$ORIG" | cut -f2-)
  R=$(printf '%s' "$RESTORED" | sed 's/^RESTORED//' | cut -f2-)
  if [ "$O" = "$R" ]; then echo "RESTORE_MATCH"; else echo "RESTORE_MISMATCH orig=[$O] restored=[$R]"; fi
fi
# ---- view-mode restore ----
if [ -n "$ORIG_VIEW" ] && [ "$ORIG_VIEW" != "inline" ] && [ "$ORIG_VIEW" != "unknown" ]; then
  runwd 90 osascript "$D/pdf_viewdoc.applescript" "$C/pdfpilot-tc-p0-del-para.docx" "$ORIG_VIEW" 2>&1
  echo "VIEW_RESTORE_RC=$?"
fi
# ---- clean x3 ----
for id in $IDS; do export_one clean "$id" word-clean.pdf; done
# ---- final state ----
echo "--- final doc state"
doc_names
echo "--- input hashes (source, after)"
shasum -a 256 "$P"/tc-*.docx | sed "s|$P/||"
echo "--- input hashes (staged copies, after)"
runwd 30 shasum -a 256 "$C"/pdfpilot-tc-*.docx | sed "s|$C/||"
echo "--- container PDF cleanup"
runwd 20 rm -f "$C"/pdfpilot-tc-*-word-*.pdf && echo "CONTAINER_PDF_CLEAN_OK" || echo "CONTAINER_PDF_CLEAN_FAIL"
echo "--- outputs"
ls -la "$P"/out/*/word-clean.pdf "$P"/out/*/word-markup-custom.pdf 2>&1
# ---- release ----
kill $HB 2>/dev/null
if wl_release_plain; then
  RELEASED=1; echo "LOCK_RELEASED plain"
else
  NAMES=$(doc_names)
  echo "PLAIN_RELEASE_REFUSED docs: $NAMES"
  case "$NAMES" in
    *pdfpilot*) echo "RELEASE_DEFERRED: pilot documents still open" ;;
    *) bash "$D/word_lock.sh" release --force --token "$TOK" >/dev/null 2>&1 && { RELEASED=1; echo "LOCK_RELEASED force (foreign documents only)"; } || echo "RELEASE_FAILED" ;;
  esac
fi
kill $CAF 2>/dev/null
echo "=== PDF_ORACLE5_DONE at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
