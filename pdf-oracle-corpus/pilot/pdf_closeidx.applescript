-- pdf_closeidx.applescript — close document 1 ONLY if its name starts
-- with "pdfpilot-" (the pilot's own staged copy). Saved-flag pattern:
-- mark saved, close with no saving argument; fall back to saving no.
with timeout of 30 seconds
	tell application "Microsoft Word"
		if (count of documents) is 0 then return "NO_DOCS"
		set tn to name of document 1
		if tn does not start with "pdfpilot-" then return "REFUSE_FOREIGN " & tn
		try
			set saved of document 1 to true
		end try
		try
			close document 1
			return "CLOSED " & tn & " docs=" & (count of documents)
		on error errMsg number errNum
			try
				close document 1 saving no
				return "CLOSED_SAVINGNO " & tn & " docs=" & (count of documents)
			on error errMsg2 number errNum2
				return "CLOSE_FAIL " & tn & " " & errNum2 & " " & errMsg2 & " docs=" & (count of documents)
			end try
		end try
	end tell
end timeout
