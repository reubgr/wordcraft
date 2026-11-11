-- pdf_axinspect.applescript — READ-ONLY census of Word's windows and,
-- for any window named "Grant File Access", its buttons (names +
-- descriptions + static texts). Clicks nothing. Used to identify the
-- exact Cancel control on the grant panels before dismissing them
-- (grant panels are never approved — see the oracle skill).
with timeout of 60 seconds
	set out to {}
	tell application "System Events"
		tell process "Microsoft Word"
			set end of out to "WINDOWS " & (count of windows)
			repeat with i from 1 to (count of windows)
				set wn to name of window i
				set end of out to ("WIN" & i & " [" & wn & "]")
				if wn is "Grant File Access" then
					try
						repeat with b in (every button of window i)
							set end of out to ("  BTN [" & (name of b) & "] desc=[" & (description of b) & "]")
						end repeat
					on error errMsg number errNum
						set end of out to ("  BTN_ENUM_FAIL " & errNum)
					end try
					try
						repeat with t in (every static text of window i)
							set end of out to ("  TXT [" & (value of t) & "]")
						end repeat
					end try
					try
						repeat with g in (every group of window i)
							repeat with b in (every button of g)
								set end of out to ("  GRPBTN [" & (name of b) & "]")
							end repeat
							repeat with t in (every static text of g)
								set end of out to ("  GRPTXT [" & (value of t) & "]")
							end repeat
						end repeat
					end try
				end if
			end repeat
		end tell
	end tell
	set AppleScript's text item delimiters to linefeed
	return out as text
end timeout
