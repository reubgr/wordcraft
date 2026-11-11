-- pdf_viewdoc.applescript — open a staged copy, set the view revisions
-- mode, close without saving (saved-flag fallback). Used by the v5
-- driver to restore the pre-custom view mode after the custom exports
-- set in-line revisions. argv: <docxPath> <balloon|inline|mixed>
on run argv
	set docxPath to item 1 of argv
	set nm to item 2 of argv
	set out to {}
	with timeout of 60 seconds
		try
			tell application "Microsoft Word"
				open docxPath
				if nm is "balloon" then set revisions mode of view of active window to balloon revisions
				if nm is "inline" then set revisions mode of view of active window to in line revisions
				if nm is "mixed" then set revisions mode of view of active window to mixed revisions
				set m to revisions mode of view of active window
				set r to "unknown"
				if m = balloon revisions then set r to "balloon"
				if m = in line revisions then set r to "inline"
				if m = mixed revisions then set r to "mixed"
				set end of out to "VIEW_SET_TO " & r
			end tell
		on error errMsg number errNum
			set end of out to ("VIEW_SET_FAIL " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to close active document saving no
			set end of out to "CLOSE_OK plain"
		on error errMsg number errNum
			try
				tell application "Microsoft Word"
					set saved of active document to true
					close active document
				end tell
				set end of out to "CLOSE_OK savedflag"
			on error errMsg2 number errNum2
				set end of out to ("CLOSE_FAIL " & errNum2 & " " & errMsg2)
			end try
		end try
		try
			tell application "Microsoft Word" to set end of out to ("FINAL_DOCS " & (count of documents))
		end try
	end timeout
	set AppleScript's text item delimiters to linefeed
	return out as text
end run
