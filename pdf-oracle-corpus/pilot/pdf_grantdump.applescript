-- pdf_grantdump.applescript — dump roles/titles/values of the FIRST
-- "Grant File Access" window via entire-contents (fetched and consumed
-- inside one tell, per the dlg_prompt v1 pattern). Read-only.
set out to {}
tell application "System Events"
	tell process "Microsoft Word"
		repeat with i from 1 to (count of windows)
			set w to window i
			set wn to ""
			try
				set wn to name of w
			end try
			if wn is "Grant File Access" then
				set end of out to "GRANT_CTX window=" & i
				try
					set welems to entire contents of w
					repeat with e in welems
						try
							set r to role of e
							if r is "AXButton" then
								set end of out to "BTN [" & (title of e) & "]"
							else if r is "AXStaticText" then
								set v to ""
								try
									set v to value of e
								end try
								if v is missing value then set v to ""
								set end of out to "TEXT [" & (v as text) & "]"
							else if r is "AXSheet" then
								set end of out to "ELEM AXSheet"
							end if
						end try
					end repeat
				on error errMsg number errNum
					set end of out to ("ENTIRE_FAIL " & errNum & " " & errMsg)
				end try
				exit repeat
			end if
		end repeat
	end tell
end tell
if (count of out) is 0 then return "NO_GRANT_WINDOW"
set AppleScript's text item delimiters to linefeed
return out as text
