-- pdf_save_test.applescript — open staged copy, show/print revisions,
-- save as PDF to container path AND out-dir path, close without saving.
with timeout of 120 seconds
	set out to {}
	tell application "Microsoft Word"
		open "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/pdfpilot-tc-p2m-multiauthor.docx"
		set doc to active document
		set end of out to ("OPENED " & (name of doc))
		set show revisions of doc to true
		set print revisions of doc to true
	end tell
	try
		tell application "Microsoft Word" to save as doc file name "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/probe-settings.pdf" file format format PDF
		set end of out to "SAVE_CONTAINER_OK"
	on error errMsg number errNum
		set end of out to ("SAVE_CONTAINER_FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to save as doc file name "/Users/muse/word-oracle/pdf-pilot/out/tc-p2m-multiauthor/probe-settings.pdf" file format format PDF
		set end of out to "SAVE_OUT_OK"
	on error errMsg number errNum
		set end of out to ("SAVE_OUT_FAIL " & errNum & " " & errMsg)
	end try
	tell application "Microsoft Word"
		close doc saving no
		set end of out to ("DONE docs=" & (count of documents))
	end tell
	set AppleScript's text item delimiters to linefeed
	return out as text
end timeout
