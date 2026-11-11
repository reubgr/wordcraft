#!/bin/bash
# pdf_axclose_run.command — recorded fallback launch (queue-runner busy
# behind another lane's lock-wait): runs pdf_axclose.applescript in the
# GUI session via LaunchServices `open`, to AX-close the pilot's wedged
# document window and answer any Don't Save sheet. Never clicks a Grant
# File Access panel (the script only clicks window close buttons and
# sheet buttons whose name contains "Don").
D=/Users/muse/compare-vs-word/dialog
OUT=$D/pdf_axclose_run.out
: > "$OUT"
exec >> "$OUT" 2>&1
echo "=== AXCLOSE_RUN at $(date -u +%Y-%m-%dT%H:%M:%SZ)"
TOK=$(cat "$HOME/word-oracle/WORD-LOCK.d/token" 2>/dev/null)
[ -n "$TOK" ] && bash "$D/word_lock.sh" heartbeat --token "$TOK" >/dev/null 2>&1 && echo HB_OK
osascript "$D/pdf_axclose.applescript" 2>&1
sleep 2
osascript -e 'with timeout of 5 seconds
tell application "Microsoft Word" to return "COUNT " & (count of documents)
end timeout' 2>&1
echo "=== AXCLOSE_RUN_DONE"
