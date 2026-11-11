-- pdf_closeone.applescript — close the FIRST open document whose name
-- starts with "pdfpilot-" (the pilot's own staged copies only; foreign
-- documents are never touched), without saving: mark it saved first
-- (showin pattern), then close with no saving argument; fall back to
-- close saving no. Prints what it did and the remaining count.
with timeout of 30 seconds
	tell application "Microsoft Word"
		set target to missing value
		repeat with d in documents
			if (name of d) starts with "pdfpilot-" then
				set target to d
				exit repeat
			end if
		end repeat
		if target is missing value then return "NO_PILOT_DOC docs=" & (count of documents)
		set tn to name of target
		try
			set saved of target to true
		end try
		try
			close target
			return "CLOSED " & tn & " docs=" & (count of documents)
		on error errMsg number errNum
			try
				close target saving no
				return "CLOSED_SAVINGNO " & tn & " docs=" & (count of documents)
			on error errMsg2 number errNum2
				return "CLOSE_FAIL " & tn & " " & errNum2 & " " & errMsg2 & " docs=" & (count of documents)
			end try
		end try
	end tell
end timeout
