-- pdf_axesc.applescript — dismiss stacked "Grant File Access" panels
-- by CANCELLING them with the Escape key (never approves a grant).
-- Only window-name reads (proven fast while panels are up); one Escape
-- per pass, census between passes, hard stop after 6 passes or when no
-- grant window remains. Ends with a final census + document count.
set out to {}
tell application "System Events"
	tell process "Microsoft Word"
		repeat with pass from 1 to 6
			set wn to {}
			repeat with i from 1 to (count of windows)
				set end of wn to (name of window i)
			end repeat
			set AppleScript's text item delimiters to " | "
			set end of out to ("PASS" & pass & " " & (wn as text))
			set AppleScript's text item delimiters to ""
			set hasGrant to false
			repeat with n in wn
				if (n as text) is "Grant File Access" then set hasGrant to true
			end repeat
			if not hasGrant then exit repeat
			key code 53
			set end of out to ("ESC_SENT pass" & pass)
			delay 2
		end repeat
		set wn to {}
		repeat with i from 1 to (count of windows)
			set end of wn to (name of window i)
		end repeat
		set AppleScript's text item delimiters to " | "
		set end of out to ("FINAL " & (wn as text))
	end tell
end tell
try
	tell application "Microsoft Word" to set end of out to ("DOCS " & (count of documents))
end try
set AppleScript's text item delimiters to linefeed
return out as text
