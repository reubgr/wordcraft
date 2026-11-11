-- pdf_settings_form.applescript — get/set Track Changes colours via the
-- application `settings` property (wSet), with equality readback and
-- self-restore. Prints one result line per formulation.
with timeout of 60 seconds
	set out to {}
	try
		tell application "Microsoft Word"
			set o to settings
			set r to "GET_VIA_SETTINGS "
			if (inserted text color of o) = blue then set r to r & "insColor=blue"
			if (inserted text color of o) = by author then set r to r & "insColor=byAuthor"
			if (inserted text color of o) = auto then set r to r & "insColor=auto"
			if (deleted text color of o) = red then set r to r & " delColor=red"
			if (deleted text color of o) = by author then set r to r & " delColor=byAuthor"
			set end of out to r
		end tell
	on error errMsg number errNum
		set end of out to ("GET_VIA_SETTINGS_FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word"
			set beforeBA to (inserted text color of settings) = by author
			set inserted text color of settings to blue
			set afterBlue to (inserted text color of settings) = blue
			set inserted text color of settings to by author
			set restored to (inserted text color of settings) = by author
			set end of out to ("SET_VIA_SETTINGS beforeByAuthor=" & beforeBA & " afterBlue=" & afterBlue & " restoredByAuthor=" & restored)
		end tell
	on error errMsg number errNum
		set end of out to ("SET_VIA_SETTINGS_FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word"
			set dmBefore to "?"
			if (deleted text mark of settings) = deleted text mark strike through then set dmBefore to "strikethrough"
			if (deleted text mark of settings) = deleted text mark hidden then set dmBefore to "hidden"
			set deleted text mark of settings to deleted text mark strike through
			set dmAfter to "?"
			if (deleted text mark of settings) = deleted text mark strike through then set dmAfter to "strikethrough"
			set end of out to ("DELMARK_VIA_SETTINGS before=" & dmBefore & " after=" & dmAfter)
		end tell
	on error errMsg number errNum
		set end of out to ("DELMARK_VIA_SETTINGS_FAIL " & errNum & " " & errMsg)
	end try
	set AppleScript's text item delimiters to linefeed
	return out as text
end timeout
