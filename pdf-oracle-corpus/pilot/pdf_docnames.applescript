-- pdf_docnames.applescript — list open Word documents (count + names).
with timeout of 20 seconds
	tell application "Microsoft Word"
		set out to "DOCS=" & (count of documents)
		repeat with d in documents
			set out to out & " | " & (name of d)
		end repeat
		return out
	end tell
end timeout
