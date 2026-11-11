-- pdf_docnames2.applescript — count + per-index names of open Word
-- documents. Index-based reads only (the repeat-over-documents
-- formulation failed with -1708 in this session; per-index name reads
-- are proven from both SSH and GUI contexts).
with timeout of 30 seconds
	tell application "Microsoft Word"
		set n to count of documents
		set out to "DOCS=" & n
		repeat with i from 1 to n
			try
				set out to out & " | " & (name of document i)
			on error errMsg number errNum
				set out to out & " | ERR" & errNum
			end try
		end repeat
		return out
	end tell
end timeout
