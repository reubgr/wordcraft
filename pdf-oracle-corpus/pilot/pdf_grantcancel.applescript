-- pdf_grantcancel.applescript — dismiss stacked "Grant File Access"
-- panels by clicking their "Cancel" button via DIRECT addressing
-- (button "Cancel" of window i — no tree traversal; traversal parks
-- behind these modal panels). NEVER clicks "Select..." (the grant
-- approval). One cancel per pass, name census between passes.
set out to {}
tell application "System Events"
	tell process "Microsoft Word"
		repeat with pass from 1 to 6
			set gi to 0
			repeat with i from 1 to (count of windows)
				if (name of window i) is "Grant File Access" then
					set gi to i
					exit repeat
				end if
			end repeat
			set end of out to ("PASS" & pass & " grantWindowIndex=" & gi)
			if gi is 0 then exit repeat
			try
				click button "Cancel" of window gi
				set end of out to ("CANCEL_CLICKED pass" & pass)
			on error errMsg number errNum
				set end of out to ("CANCEL_FAIL pass" & pass & " " & errNum & " " & errMsg)
				exit repeat
			end try
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
