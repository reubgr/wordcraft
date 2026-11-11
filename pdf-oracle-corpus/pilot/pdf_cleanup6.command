#!/bin/bash
# pdf_cleanup6.command — v2 (rewritten before the queued job started).
# Recovery after the killed v6 driver (#6153): adopts the dead driver's
# OWN claim (token read from the lock dir on the mini — the claim is
# this pilot's), censuses windows traversal-free, closes the pilot's
# wedged document via the bundle showin_close pattern ONLY if every
# open document is positively a pdfpilot-* staged copy, deletes the
# wrong-document PDFs the poisoned run produced (multiauthor +
# table-row-del markup-custom, all clean files, all container PDFs),
# confirms options by snapshot, and releases. Force-release ONLY when a
# positive census shows no pdfpilot document remains (foreign docs are
# never closed by this script).
D=/Users/muse/compare-vs-word/dialog
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
OUT=$D/pdf_cleanup6.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== CLEANUP6 v2 at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
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
TOK=$(cat "$HOME/word-oracle/WORD-LOCK.d/token" 2>/dev/null)
if [ -z "$TOK" ]; then echo "NO_TOKEN_IN_LOCKDIR"; echo "=== CLEANUP6_DONE rc=notoken"; exit 3; fi
bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 && echo "CLAIM_ADOPTED heartbeat ok" || echo "HEARTBEAT_FAILED"
( while :; do sleep 45; bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 || break; done ) &
HB=$!
echo "--- window census (traversal-free)"
runwd 30 osascript "$D/dlg_prompt.applescript" census 2>&1
echo "--- docs before"
NAMES=$(runwd 40 osascript "$D/pdf_docnames2.applescript" 2>&1)
echo "$NAMES"
# Close gate: showin_close closes EVERY document, so run it only when
# the census positively shows pdfpilot docs and nothing foreign.
FOREIGN=$(printf '%s' "$NAMES" | grep -o 'DOCS=[0-9]*' )
case "$NAMES" in
  *pdfpilot*DOCS*|*DOCS=0*) : ;;
esac
if printf '%s' "$NAMES" | grep -q "pdfpilot-"; then
  # every listed name must be pdfpilot-* or the DOCS= prefix itself
  REST=$(printf '%s' "$NAMES" | sed 's/^DOCS=[0-9]*//' | tr '|' '\n' | grep -v '^ *$' | grep -vc 'pdfpilot-')
  if [ "$REST" = "0" ]; then
    echo "--- closing via showin_close (all open docs are pilot docs)"
    runwd 60 osascript "$D/showin_close.applescript" 2>&1
  else
    echo "CLOSE_SKIPPED: foreign documents present ($NAMES)"
  fi
else
  echo "CLOSE_SKIPPED: no pdfpilot document in census"
fi
echo "--- docs after"
NAMES2=$(runwd 40 osascript "$D/pdf_docnames2.applescript" 2>&1)
echo "$NAMES2"
echo "--- window census after"
runwd 30 osascript "$D/dlg_prompt.applescript" census 2>&1
echo "--- wrong-output cleanup"
rm -f "$P/out/tc-p2m-multiauthor/word-markup-custom.pdf" "$P/out/tc-p2m-table-row-del/word-markup-custom.pdf"
rm -f "$P"/out/*/word-clean.pdf
runwd 20 rm -f "$C"/pdfpilot-tc-*-word-*.pdf && echo "CONTAINER_PDFS_REMOVED" || echo "CONTAINER_PDF_RM_FAIL"
ls -la "$P"/out/*/ 2>&1
echo "--- options snapshot (expect restored originals)"
runwd 90 osascript "$D/pdf_opts6.applescript" snapshot 2>&1
echo "--- release"
kill $HB 2>/dev/null
if bash "$D/word_lock.sh" release --token "$TOK" >/dev/null 2>&1; then
  echo "LOCK_RELEASED plain"
else
  echo "PLAIN_RELEASE_REFUSED; final census: $NAMES2"
  case "$NAMES2" in
    *pdfpilot*) echo "RELEASE_DEFERRED: pilot documents still open" ;;
    *DOCS=*) bash "$D/word_lock.sh" release --force --token "$TOK" >/dev/null 2>&1 && echo "LOCK_RELEASED force (no pilot documents remain)" || echo "RELEASE_FAILED" ;;
    *) echo "RELEASE_DEFERRED: census unreadable, not forcing" ;;
  esac
fi
echo "=== CLEANUP6_DONE"
