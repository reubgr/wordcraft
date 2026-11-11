#!/bin/bash
# probe_tcc.sh — diagnosis only: is a TCC prompt up in the GUI session, and
# does a shell cp into Word's container still block? No Word driving, no lock.
C="$HOME/Library/Containers/com.microsoft.Word/Data/Documents"
P="$HOME/word-oracle/pdf-pilot"
echo "PROBE_START $(date -u +%Y-%m-%dT%H:%M:%SZ)"
osascript <<'APPLESCRIPT' 2>&1
tell application "System Events"
	set out to {}
	repeat with pr in (every process whose background only is false)
		try
			set wc to count of windows of pr
			if wc > 0 then
				set wn to {}
				repeat with i from 1 to wc
					set end of wn to (name of window i of pr)
				end repeat
				set end of out to ((name of pr) & " wins=" & wc & " [" & (wn as text) & "]")
			end if
		end try
	end repeat
	return out
end tell
APPLESCRIPT
echo "--- cp test (20s watchdog)"
cp "$P/tc-p0-del-para.docx" "$C/probe_tcc_test.docx" &
CPPID=$!
for i in $(seq 1 20); do
  if ! kill -0 $CPPID 2>/dev/null; then break; fi
  sleep 1
done
if kill -0 $CPPID 2>/dev/null; then
  kill -9 $CPPID 2>/dev/null
  echo "CP_IN_BLOCKED after 20s"
else
  wait $CPPID; echo "CP_IN_OK rc=$?"
fi
ls -la "$C/probe_tcc_test.docx" 2>&1 | head -2
rm -f "$C/probe_tcc_test.docx" 2>/dev/null
echo "PROBE_DONE"
