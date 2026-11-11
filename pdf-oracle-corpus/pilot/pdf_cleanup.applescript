-- pdf_cleanup.applescript — report open Word documents, close all WITHOUT saving.
with timeout of 60 seconds
	tell application "Microsoft Word"
		set out to "BEFORE docs=" & (count of documents)
		repeat with d in documents
			set out to out & " | " & (name of d)
		end repeat
		repeat while (count of documents) > 0
			close document 1 saving no
		end repeat
		set out to out & " AFTER docs=" & (count of documents)
		return out
	end tell
end timeout
