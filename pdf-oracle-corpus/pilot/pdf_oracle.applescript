-- pdf_oracle.applescript v2 — Word->PDF oracle pilot (storytold/wordcraft#6153).
-- Runs inside the muse console session (queue-runner job) where Terminal's
-- grants apply. Job lines are EMBEDDED (v1 read them from a TSV in Word's
-- container via StandardAdditions and the read failed silently — ERRTOP;
-- embedded lines remove that file read entirely). For each job: opens the
-- staged container copy, sets the display/print revision state for the
-- mode, saves as PDF (dictionary route: save as, file format format PDF),
-- closes WITHOUT saving. Appends results + full Track Changes Options
-- snapshots to pdf_log.txt in the container. Original Word options are
-- restored at the end. Order matters: all markup-default jobs run before
-- any markup-custom job touches the options.
with timeout of 1800 seconds
	set C to "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/"
	set logPath to C & "pdf_log.txt"
	my logLine(logPath, "PDF_ORACLE_START v2")
	set jobs to {{"markup-default", "tc-p0-del-para"}, {"markup-default", "tc-p2m-multiauthor"}, {"markup-default", "tc-p2m-table-row-del"}, {"markup-custom", "tc-p0-del-para"}, {"markup-custom", "tc-p2m-multiauthor"}, {"markup-custom", "tc-p2m-table-row-del"}, {"clean", "tc-p0-del-para"}, {"clean", "tc-p2m-multiauthor"}, {"clean", "tc-p2m-table-row-del"}}
	set origOpts to my snapshotOptions("ORIGINAL")
	my logLine(logPath, origOpts)
	repeat with j in jobs
		set mode to item 1 of j
		set eid to item 2 of j
		set inPath to C & "pdfpilot-" & eid & ".docx"
		set outPath to C & "pdfpilot-" & eid & "-word-" & mode & ".pdf"
		my oneJob(logPath, mode, inPath, outPath)
	end repeat
	my restoreOptions(origOpts, logPath)
	my logLine(logPath, my snapshotOptions("RESTORED"))
	my logLine(logPath, "PDF_ORACLE_DONE v2")
end timeout

on oneJob(logPath, mode, inPath, outPath)
	set t0 to (current date)
	try
		tell application "Microsoft Word"
			open inPath
			set doc to active document
			if mode is "markup-custom" then
				set inserted text color of Word options to blue
				set inserted text mark of Word options to inserted text mark underline
				set deleted text color of Word options to red
				set deleted text mark of Word options to deleted text mark strike through
			end if
			if mode is "clean" then
				set show revisions of doc to false
				set print revisions of doc to false
			else
				set show revisions of doc to true
				set print revisions of doc to true
			end if
			try
				set revisions view of view of active window to revisions view final
			end try
		end tell
		my logLine(logPath, my snapshotOptions("FOR_" & mode))
		tell application "Microsoft Word"
			save as doc file name outPath file format format PDF
		end tell
		try
			tell application "Microsoft Word" to close doc saving no
		on error
			tell application "Microsoft Word" to close document 1 saving no
		end try
		set secs to ((current date) - t0) as integer
		my logLine(logPath, "OK" & tab & mode & tab & outPath & tab & secs)
	on error errMsg number errNum
		my logLine(logPath, "ERR" & tab & mode & tab & outPath & tab & errNum & " " & errMsg)
		try
			tell application "Microsoft Word"
				repeat while (count of documents) > 0
					close document 1 saving no
				end repeat
			end tell
		end try
	end try
end oneJob

on snapshotOptions(tag)
	try
		tell application "Microsoft Word"
			set o to Word options
			set s to tag
			set s to s & tab & "insColor=" & (inserted text color of o as text)
			set s to s & tab & "insMark=" & (inserted text mark of o as text)
			set s to s & tab & "delColor=" & (deleted text color of o as text)
			set s to s & tab & "delMark=" & (deleted text mark of o as text)
			set s to s & tab & "insCellColor=" & (inserted cell color of o as text)
			set s to s & tab & "delCellColor=" & (deleted cell color of o as text)
			set s to s & tab & "mergedCellColor=" & (merged cell color of o as text)
			set s to s & tab & "splitCellColor=" & (split cell color of o as text)
			set s to s & tab & "moveFromColor=" & (move from text color of o as text)
			set s to s & tab & "moveFromMark=" & (move from text mark of o as text)
			set s to s & tab & "moveToColor=" & (move to text color of o as text)
			set s to s & tab & "moveToMark=" & (move to text mark of o as text)
			set s to s & tab & "commentsColor=" & (comments color of o as text)
			set s to s & tab & "printComments=" & (print comments of o as text)
			return s
		end tell
	on error errMsg number errNum
		return tag & tab & "SNAPSHOT_FAIL " & errNum & " " & errMsg
	end try
end snapshotOptions

on restoreOptions(origLine, logPath)
	try
		set AppleScript's text item delimiters to tab
		set f to text items of origLine
		set AppleScript's text item delimiters to ""
		set vals to {}
		repeat with kv in f
			set kvStr to kv as text
			if kvStr contains "=" then
				set AppleScript's text item delimiters to "="
				set end of vals to item 2 of (text items of kvStr)
				set AppleScript's text item delimiters to ""
			end if
		end repeat
		-- order in snapshot: insColor, insMark, delColor, delMark, insCell,
		-- delCell, mergedCell, splitCell, moveFromColor, moveFromMark,
		-- moveToColor, moveToMark, commentsColor, printComments
		if (count of vals) = 14 then
			my applyOptions(item 1 of vals, item 2 of vals, item 3 of vals, item 4 of vals, item 9 of vals, item 10 of vals, item 11 of vals, item 12 of vals)
		else
			my logLine(logPath, "RESTORE_SKIP vals=" & (count of vals))
		end if
	on error errMsg number errNum
		my logLine(logPath, "RESTORE_FAIL " & errNum & " " & errMsg)
	end try
end restoreOptions

-- All enum literals live inside the tell block: multi-word enumerators
-- (dark blue, by author, ...) do not compile outside one (-2741).
on applyOptions(ic, im, dc, dm, mfc, mfm, mtc, mtm)
	tell application "Microsoft Word"
		set o to Word options
		if ic is "blue" then set inserted text color of o to blue
		if ic is "red" then set inserted text color of o to red
		if ic is "black" then set inserted text color of o to black
		if ic is "auto" then set inserted text color of o to auto
		if ic is "by author" then set inserted text color of o to by author
		if ic is "pink" then set inserted text color of o to pink
		if ic is "green" then set inserted text color of o to green
		if ic is "violet" then set inserted text color of o to violet
		if ic is "yellow" then set inserted text color of o to yellow
		if ic is "white" then set inserted text color of o to white
		if ic is "dark blue" then set inserted text color of o to dark blue
		if ic is "dark red" then set inserted text color of o to dark red
		if ic is "turquoise" then set inserted text color of o to turquoise
		if ic is "bright green" then set inserted text color of o to bright green
		if ic is "teal" then set inserted text color of o to teal
		if ic is "dark yellow" then set inserted text color of o to dark yellow
		if ic is "gray50" then set inserted text color of o to gray50
		if ic is "gray25" then set inserted text color of o to gray25
		if im contains "double underline" then set inserted text mark of o to inserted text mark double underline
		if im contains "underline" and im does not contain "double" then set inserted text mark of o to inserted text mark underline
		if im contains "bold" then set inserted text mark of o to inserted text mark bold
		if im contains "italic" then set inserted text mark of o to inserted text mark italic
		if im contains "color only" then set inserted text mark of o to inserted text mark color only
		if im contains "double strike" then set inserted text mark of o to inserted text mark double strike through
		if im contains "strike" and im does not contain "double" then set inserted text mark of o to inserted text mark strike through
		if im contains "none" then set inserted text mark of o to inserted text mark none
		if dc is "blue" then set deleted text color of o to blue
		if dc is "red" then set deleted text color of o to red
		if dc is "black" then set deleted text color of o to black
		if dc is "auto" then set deleted text color of o to auto
		if dc is "by author" then set deleted text color of o to by author
		if dc is "pink" then set deleted text color of o to pink
		if dc is "green" then set deleted text color of o to green
		if dc is "violet" then set deleted text color of o to violet
		if dc is "yellow" then set deleted text color of o to yellow
		if dc is "white" then set deleted text color of o to white
		if dc is "dark blue" then set deleted text color of o to dark blue
		if dc is "dark red" then set deleted text color of o to dark red
		if dc is "turquoise" then set deleted text color of o to turquoise
		if dc is "bright green" then set deleted text color of o to bright green
		if dc is "teal" then set deleted text color of o to teal
		if dc is "dark yellow" then set deleted text color of o to dark yellow
		if dc is "gray50" then set deleted text color of o to gray50
		if dc is "gray25" then set deleted text color of o to gray25
		if dm contains "double strike" then set deleted text mark of o to deleted text mark double strike through
		if dm contains "strike" and dm does not contain "double" then set deleted text mark of o to deleted text mark strike through
		if dm contains "caret" then set deleted text mark of o to deleted text mark caret
		if dm contains "pound" then set deleted text mark of o to deleted text mark pound
		if dm contains "bold" then set deleted text mark of o to deleted text mark bold
		if dm contains "italic" then set deleted text mark of o to deleted text mark italic
		if dm contains "underline" then set deleted text mark of o to deleted text mark underline
		if dm contains "color only" then set deleted text mark of o to deleted text mark color only
		if dm contains "hidden" then set deleted text mark of o to deleted text mark hidden
		if dm contains "none" then set deleted text mark of o to deleted text mark none
		if mfc is "blue" then set move from text color of o to blue
		if mfc is "red" then set move from text color of o to red
		if mfc is "black" then set move from text color of o to black
		if mfc is "auto" then set move from text color of o to auto
		if mfc is "by author" then set move from text color of o to by author
		if mfc is "pink" then set move from text color of o to pink
		if mfc is "green" then set move from text color of o to green
		if mfc is "violet" then set move from text color of o to violet
		if mfc is "dark blue" then set move from text color of o to dark blue
		if mfc is "dark red" then set move from text color of o to dark red
		if mfc is "bright green" then set move from text color of o to bright green
		if mfm contains "double strike" then set move from text mark of o to deleted text mark double strike through
		if mfm contains "strike" and mfm does not contain "double" then set move from text mark of o to deleted text mark strike through
		if mfm contains "caret" then set move from text mark of o to deleted text mark caret
		if mfm contains "hidden" then set move from text mark of o to deleted text mark hidden
		if mfm contains "none" then set move from text mark of o to deleted text mark none
		if mtc is "blue" then set move to text color of o to blue
		if mtc is "red" then set move to text color of o to red
		if mtc is "black" then set move to text color of o to black
		if mtc is "auto" then set move to text color of o to auto
		if mtc is "by author" then set move to text color of o to by author
		if mtc is "pink" then set move to text color of o to pink
		if mtc is "green" then set move to text color of o to green
		if mtc is "violet" then set move to text color of o to violet
		if mtc is "dark blue" then set move to text color of o to dark blue
		if mtc is "dark red" then set move to text color of o to dark red
		if mtc is "bright green" then set move to text color of o to bright green
		if mtm contains "double underline" then set move to text mark of o to inserted text mark double underline
		if mtm contains "underline" and mtm does not contain "double" then set move to text mark of o to inserted text mark underline
		if mtm contains "bold" then set move to text mark of o to inserted text mark bold
		if mtm contains "italic" then set move to text mark of o to inserted text mark italic
		if mtm contains "color only" then set move to text mark of o to inserted text mark color only
		if mtm contains "none" then set move to text mark of o to inserted text mark none
	end tell
end applyOptions

on logLine(logPath, s)
	try
		set fd to open for access (POSIX file logPath) with write permission
		write (s & linefeed) to fd starting at eof
		close access fd
	on error
		try
			close access (POSIX file logPath)
		end try
	end try
end logLine
