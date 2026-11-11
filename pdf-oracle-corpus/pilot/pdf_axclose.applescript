-- pdf_axclose.applescript — AX close of the open Word document:
-- click window 1's close button, answer any save prompt with the
-- Don't Save button, report the final document count.
with timeout of 60 seconds
	tell application "System Events"
		tell process "Microsoft Word"
			set wn to {}
			repeat with i from 1 to (count of windows)
				set end of wn to ("WIN" & i & " [" & (name of window i) & "]")
			end repeat
			set AppleScript's text item delimiters to " | "
			log (wn as text)
			if (count of windows) > 0 then
				try
					click button 1 of window 1
					log "CLICKED close button of window 1"
				on error errMsg number errNum
					log ("CLOSE_BTN_FAIL " & errNum & " " & errMsg)
				end try
			end if
		end tell
	end tell
	delay 2
	tell application "System Events"
		tell process "Microsoft Word"
			repeat with i from 1 to (count of windows)
				try
					repeat with b in (every button of sheet 1 of window i)
						if (name of b) contains "Don" then
							click b
							log ("CLICKED prompt button " & (name of b))
						end if
					end repeat
				end try
			end repeat
		end tell
	end tell
	delay 2
	tell application "Microsoft Word" to return "docs=" & (count of documents)
end timeout
