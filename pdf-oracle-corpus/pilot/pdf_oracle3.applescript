-- pdf_oracle3.applescript — Word->PDF oracle pilot RESUMPTION (#6153).
-- Continues pilot-report.md: markup-default x3 already produced in run 3.
-- This run produces markup-custom x3 and clean x3 for the same entries,
-- and completes the Track Changes Options readback that run 3 could not
-- do (its `as text` coercion of colour properties failed with -1700).
-- Readback here is EQUALITY-based against dictionary constants, per the
-- withdrawn probe's design. All enum literals live inside tell blocks
-- (multi-word enumerators do not compile outside one, -2741).
-- Order: snapshot ORIGINAL -> isolated custom sets -> snapshot CUSTOM ->
-- custom exports -> restore originals -> snapshot RESTORED -> view-mode
-- restore -> clean exports. Documents are closed WITHOUT saving.
with timeout of 1800 seconds
	set C to "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/"
	set logPath to C & "pdf_log3.txt"
	my logLine(logPath, "PDF_ORACLE3_START")
	set origOpts to my snapshotOptions("ORIGINAL")
	my logLine(logPath, origOpts)
	-- isolated custom sets, each logged on its own
	try
		tell application "Microsoft Word" to set inserted text color of Word options to blue
		my logLine(logPath, "SET_OK insColor=blue")
	on error errMsg number errNum
		my logLine(logPath, "SET_FAIL insColor " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set inserted text mark of Word options to inserted text mark underline
		my logLine(logPath, "SET_OK insMark=underline")
	on error errMsg number errNum
		my logLine(logPath, "SET_FAIL insMark " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set deleted text color of Word options to red
		my logLine(logPath, "SET_OK delColor=red")
	on error errMsg number errNum
		my logLine(logPath, "SET_FAIL delColor " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set deleted text mark of Word options to deleted text mark strike through
		my logLine(logPath, "SET_OK delMark=strikethrough")
	on error errMsg number errNum
		my logLine(logPath, "SET_FAIL delMark " & errNum & " " & errMsg)
	end try
	my logLine(logPath, my snapshotOptions("CUSTOM"))
	set ids to {"tc-p0-del-para", "tc-p2m-multiauthor", "tc-p2m-table-row-del"}
	set origViewMode to ""
	repeat with eid in ids
		set r to my exportJob(logPath, "markup-custom", eid as text, C)
		if origViewMode is "" then set origViewMode to r
	end repeat
	my logLine(logPath, "ORIG_VIEW_MODE " & origViewMode)
	-- restore original Track Changes Options
	my restoreFromSnapshot(origOpts, logPath)
	my logLine(logPath, my snapshotOptions("RESTORED"))
	-- restore the view revisions mode captured before the custom sets
	if origViewMode is not "" then
		try
			tell application "Microsoft Word"
				open C & "pdfpilot-tc-p0-del-para.docx"
				my setViewMode(origViewMode)
				close active document saving no
			end tell
			my logLine(logPath, "VIEW_RESTORE_OK " & origViewMode)
		on error errMsg number errNum
			my logLine(logPath, "VIEW_RESTORE_FAIL " & errNum & " " & errMsg)
			try
				tell application "Microsoft Word" to close document 1 saving no
			end try
		end try
	end if
	repeat with eid in ids
		my exportJob(logPath, "clean", eid as text, C)
	end repeat
	my logLine(logPath, "PDF_ORACLE3_DONE")
end timeout

-- Returns the view revisions mode string seen before any set ("" if none).
on exportJob(logPath, mode, eid, C)
	set inPath to C & "pdfpilot-" & eid & ".docx"
	set outPath to C & "pdfpilot-" & eid & "-word-" & mode & ".pdf"
	set t0 to (current date)
	set viewBefore to ""
	try
		tell application "Microsoft Word"
			open inPath
			set doc to active document
			set viewBefore to my viewModeName()
			if mode is "markup-custom" then
				set revisions mode of view of active window to in line revisions
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
		my logLine(logPath, my snapshotView("VIEW", mode, eid))
		tell application "Microsoft Word"
			save as doc file name outPath file format format PDF
		end tell
		try
			tell application "Microsoft Word" to close doc saving no
		on error
			tell application "Microsoft Word" to close document 1 saving no
		end try
		set secs to ((current date) - t0) as integer
		my logLine(logPath, "OK" & tab & mode & tab & eid & tab & secs)
	on error errMsg number errNum
		my logLine(logPath, "ERR" & tab & mode & tab & eid & tab & errNum & " " & errMsg)
		try
			tell application "Microsoft Word"
				repeat while (count of documents) > 0
					close document 1 saving no
				end repeat
			end tell
		end try
	end try
	return viewBefore
end exportJob

on viewModeName()
	tell application "Microsoft Word"
		set m to revisions mode of view of active window
		if m = balloon revisions then return "balloon"
		if m = in line revisions then return "inline"
		if m = mixed revisions then return "mixed"
		return "unknown"
	end tell
end viewModeName

on setViewMode(name)
	tell application "Microsoft Word"
		if name is "balloon" then set revisions mode of view of active window to balloon revisions
		if name is "inline" then set revisions mode of view of active window to in line revisions
		if name is "mixed" then set revisions mode of view of active window to mixed revisions
	end tell
end setViewMode

on snapshotView(tag, mode, eid)
	try
		tell application "Microsoft Word"
			set doc to active document
			set vm to my viewModeName()
			set rv to "?"
			if (revisions view of view of active window) = revisions view final then set rv to "final"
			if (revisions view of view of active window) = revisions view original then set rv to "original"
			return tag & tab & mode & tab & eid & tab & "showRev=" & (show revisions of doc as text) & tab & "printRev=" & (print revisions of doc as text) & tab & "revView=" & rv & tab & "revMode=" & vm
		end tell
	on error errMsg number errNum
		return tag & tab & mode & tab & eid & tab & "VIEW_SNAPSHOT_FAIL " & errNum & " " & errMsg
	end try
end snapshotView

on snapshotOptions(tag)
	try
		tell application "Microsoft Word"
			set o to Word options
			set s to tag
			set s to s & tab & "insColor=" & my colorName(inserted text color of o)
			set s to s & tab & "insMark=" & my insMarkName(inserted text mark of o)
			set s to s & tab & "delColor=" & my colorName(deleted text color of o)
			set s to s & tab & "delMark=" & my delMarkName(deleted text mark of o)
			set s to s & tab & "insCellColor=" & my cellColorName(inserted cell color of o)
			set s to s & tab & "delCellColor=" & my cellColorName(deleted cell color of o)
			set s to s & tab & "mergedCellColor=" & my cellColorName(merged cell color of o)
			set s to s & tab & "splitCellColor=" & my cellColorName(split cell color of o)
			set s to s & tab & "moveFromColor=" & my colorName(move from text color of o)
			set s to s & tab & "moveFromMark=" & my moveFromMarkName(move from text mark of o)
			set s to s & tab & "moveToColor=" & my colorName(move to text color of o)
			set s to s & tab & "moveToMark=" & my moveToMarkName(move to text mark of o)
			set s to s & tab & "commentsColor=" & my colorName(comments color of o)
			set s to s & tab & "printComments=" & (print comments of o as text)
			return s
		end tell
	on error errMsg number errNum
		return tag & tab & "SNAPSHOT_FAIL " & errNum & " " & errMsg
	end try
end snapshotOptions

on colorName(v)
	tell application "Microsoft Word"
		if v = auto then return "auto"
		if v = black then return "black"
		if v = blue then return "blue"
		if v = turquoise then return "turquoise"
		if v = bright green then return "brightGreen"
		if v = pink then return "pink"
		if v = red then return "red"
		if v = yellow then return "yellow"
		if v = white then return "white"
		if v = dark blue then return "darkBlue"
		if v = teal then return "teal"
		if v = green then return "green"
		if v = violet then return "violet"
		if v = dark red then return "darkRed"
		if v = dark yellow then return "darkYellow"
		if v = gray50 then return "gray50"
		if v = gray25 then return "gray25"
		if v = by author then return "byAuthor"
		return "unknown"
	end tell
end colorName

on insMarkName(v)
	tell application "Microsoft Word"
		if v = inserted text mark none then return "none"
		if v = inserted text mark bold then return "bold"
		if v = inserted text mark italic then return "italic"
		if v = inserted text mark underline then return "underline"
		if v = inserted text mark double underline then return "doubleUnderline"
		if v = inserted text mark color only then return "colorOnly"
		if v = inserted text mark strike through then return "strikethrough"
		if v = inserted text mark double strike through then return "doubleStrikethrough"
		return "unknown"
	end tell
end insMarkName

on delMarkName(v)
	tell application "Microsoft Word"
		if v = deleted text mark hidden then return "hidden"
		if v = deleted text mark strike through then return "strikethrough"
		if v = deleted text mark caret then return "caret"
		if v = deleted text mark pound then return "pound"
		if v = deleted text mark none then return "none"
		if v = deleted text mark bold then return "bold"
		if v = deleted text mark italic then return "italic"
		if v = deleted text mark underline then return "underline"
		if v = deleted text mark double underline then return "doubleUnderline"
		if v = deleted text mark color only then return "colorOnly"
		if v = deleted text mark double strike through then return "doubleStrikethrough"
		return "unknown"
	end tell
end delMarkName

on moveFromMarkName(v)
	tell application "Microsoft Word"
		if v = move from text mark hidden then return "hidden"
		if v = move from text mark double strike through then return "doubleStrikethrough"
		if v = move from text mark strike through then return "strikethrough"
		if v = move from text mark caret then return "caret"
		if v = move from text mark pound then return "pound"
		if v = move from text mark none then return "none"
		if v = move from text mark bold then return "bold"
		if v = move from text mark italic then return "italic"
		if v = move from text mark underline then return "underline"
		if v = move from text mark double underline then return "doubleUnderline"
		if v = move from text mark color only then return "colorOnly"
		return "unknown"
	end tell
end moveFromMarkName

on moveToMarkName(v)
	tell application "Microsoft Word"
		if v = move to text mark none then return "none"
		if v = move to text mark bold then return "bold"
		if v = move to text mark italic then return "italic"
		if v = move to text mark underline then return "underline"
		if v = move to text mark double underline then return "doubleUnderline"
		if v = move to text mark color only then return "colorOnly"
		if v = move to text mark strike through then return "strikethrough"
		if v = move to text mark double strike through then return "doubleStrikethrough"
		return "unknown"
	end tell
end moveToMarkName

on cellColorName(v)
	tell application "Microsoft Word"
		if v = cell color by author then return "byAuthor"
		if v = cell color no highlight then return "none"
		if v = cell color pink then return "pink"
		if v = cell color light blue then return "lightBlue"
		if v = cell color light yellow then return "lightYellow"
		if v = cell color light purple then return "lightPurple"
		if v = cell color light orange then return "lightOrange"
		if v = cell color light green then return "lightGreen"
		if v = cell color light gray then return "lightGray"
		return "unknown"
	end tell
end cellColorName

on valOf(line, key)
	set AppleScript's text item delimiters to tab
	set parts to text items of line
	set AppleScript's text item delimiters to ""
	repeat with p in parts
		set ps to p as text
		if ps starts with (key & "=") then return text ((length of key) + 2) thru -1 of ps
	end repeat
	return ""
end valOf

on restoreFromSnapshot(line, logPath)
	try
		set ic to my valOf(line, "insColor")
		set im to my valOf(line, "insMark")
		set dc to my valOf(line, "delColor")
		set dm to my valOf(line, "delMark")
		set icc to my valOf(line, "insCellColor")
		set dcc to my valOf(line, "delCellColor")
		set mcc to my valOf(line, "mergedCellColor")
		set scc to my valOf(line, "splitCellColor")
		set mfc to my valOf(line, "moveFromColor")
		set mfm to my valOf(line, "moveFromMark")
		set mtc to my valOf(line, "moveToColor")
		set mtm to my valOf(line, "moveToMark")
		set cc to my valOf(line, "commentsColor")
		my applyAll(ic, im, dc, dm, icc, dcc, mcc, scc, mfc, mfm, mtc, mtm, cc)
		my logLine(logPath, "RESTORE_APPLIED")
	on error errMsg number errNum
		my logLine(logPath, "RESTORE_FAIL " & errNum & " " & errMsg)
	end try
end restoreFromSnapshot

on applyAll(ic, im, dc, dm, icc, dcc, mcc, scc, mfc, mfm, mtc, mtm, cc)
	tell application "Microsoft Word"
		set o to Word options
		my setColorProp("ins", ic)
		if im is "none" then set inserted text mark of o to inserted text mark none
		if im is "bold" then set inserted text mark of o to inserted text mark bold
		if im is "italic" then set inserted text mark of o to inserted text mark italic
		if im is "underline" then set inserted text mark of o to inserted text mark underline
		if im is "doubleUnderline" then set inserted text mark of o to inserted text mark double underline
		if im is "colorOnly" then set inserted text mark of o to inserted text mark color only
		if im is "strikethrough" then set inserted text mark of o to inserted text mark strike through
		if im is "doubleStrikethrough" then set inserted text mark of o to inserted text mark double strike through
		my setColorProp("del", dc)
		if dm is "hidden" then set deleted text mark of o to deleted text mark hidden
		if dm is "strikethrough" then set deleted text mark of o to deleted text mark strike through
		if dm is "caret" then set deleted text mark of o to deleted text mark caret
		if dm is "pound" then set deleted text mark of o to deleted text mark pound
		if dm is "none" then set deleted text mark of o to deleted text mark none
		if dm is "bold" then set deleted text mark of o to deleted text mark bold
		if dm is "italic" then set deleted text mark of o to deleted text mark italic
		if dm is "underline" then set deleted text mark of o to deleted text mark underline
		if dm is "doubleUnderline" then set deleted text mark of o to deleted text mark double underline
		if dm is "colorOnly" then set deleted text mark of o to deleted text mark color only
		if dm is "doubleStrikethrough" then set deleted text mark of o to deleted text mark double strike through
		my setCellColorProp("ins", icc)
		my setCellColorProp("del", dcc)
		my setCellColorProp("merged", mcc)
		my setCellColorProp("split", scc)
		my setColorProp("moveFrom", mfc)
		if mfm is "hidden" then set move from text mark of o to move from text mark hidden
		if mfm is "doubleStrikethrough" then set move from text mark of o to move from text mark double strike through
		if mfm is "strikethrough" then set move from text mark of o to move from text mark strike through
		if mfm is "caret" then set move from text mark of o to move from text mark caret
		if mfm is "pound" then set move from text mark of o to move from text mark pound
		if mfm is "none" then set move from text mark of o to move from text mark none
		if mfm is "bold" then set move from text mark of o to move from text mark bold
		if mfm is "italic" then set move from text mark of o to move from text mark italic
		if mfm is "underline" then set move from text mark of o to move from text mark underline
		if mfm is "doubleUnderline" then set move from text mark of o to move from text mark double underline
		if mfm is "colorOnly" then set move from text mark of o to move from text mark color only
		my setColorProp("moveTo", mtc)
		if mtm is "none" then set move to text mark of o to move to text mark none
		if mtm is "bold" then set move to text mark of o to move to text mark bold
		if mtm is "italic" then set move to text mark of o to move to text mark italic
		if mtm is "underline" then set move to text mark of o to move to text mark underline
		if mtm is "doubleUnderline" then set move to text mark of o to move to text mark double underline
		if mtm is "colorOnly" then set move to text mark of o to move to text mark color only
		if mtm is "strikethrough" then set move to text mark of o to move to text mark strike through
		if mtm is "doubleStrikethrough" then set move to text mark of o to move to text mark double strike through
		my setColorProp("comments", cc)
	end tell
end applyAll

on setColorProp(which, nm)
	tell application "Microsoft Word"
		set o to Word options
		if nm is "auto" then my setOneColor(which, auto)
		if nm is "black" then my setOneColor(which, black)
		if nm is "blue" then my setOneColor(which, blue)
		if nm is "turquoise" then my setOneColor(which, turquoise)
		if nm is "brightGreen" then my setOneColor(which, bright green)
		if nm is "pink" then my setOneColor(which, pink)
		if nm is "red" then my setOneColor(which, red)
		if nm is "yellow" then my setOneColor(which, yellow)
		if nm is "white" then my setOneColor(which, white)
		if nm is "darkBlue" then my setOneColor(which, dark blue)
		if nm is "teal" then my setOneColor(which, teal)
		if nm is "green" then my setOneColor(which, green)
		if nm is "violet" then my setOneColor(which, violet)
		if nm is "darkRed" then my setOneColor(which, dark red)
		if nm is "darkYellow" then my setOneColor(which, dark yellow)
		if nm is "gray50" then my setOneColor(which, gray50)
		if nm is "gray25" then my setOneColor(which, gray25)
		if nm is "byAuthor" then my setOneColor(which, by author)
	end tell
end setColorProp

on setOneColor(which, v)
	tell application "Microsoft Word"
		set o to Word options
		if which is "ins" then set inserted text color of o to v
		if which is "del" then set deleted text color of o to v
		if which is "moveFrom" then set move from text color of o to v
		if which is "moveTo" then set move to text color of o to v
		if which is "comments" then set comments color of o to v
	end tell
end setOneColor

on setCellColorProp(which, nm)
	tell application "Microsoft Word"
		set o to Word options
		if nm is "byAuthor" then my setOneCellColor(which, cell color by author)
		if nm is "none" then my setOneCellColor(which, cell color no highlight)
		if nm is "pink" then my setOneCellColor(which, cell color pink)
		if nm is "lightBlue" then my setOneCellColor(which, cell color light blue)
		if nm is "lightYellow" then my setOneCellColor(which, cell color light yellow)
		if nm is "lightPurple" then my setOneCellColor(which, cell color light purple)
		if nm is "lightOrange" then my setOneCellColor(which, cell color light orange)
		if nm is "lightGreen" then my setOneCellColor(which, cell color light green)
		if nm is "lightGray" then my setOneCellColor(which, cell color light gray)
	end tell
end setCellColorProp

on setOneCellColor(which, v)
	tell application "Microsoft Word"
		set o to Word options
		if which is "ins" then set inserted cell color of o to v
		if which is "del" then set deleted cell color of o to v
		if which is "merged" then set merged cell color of o to v
		if which is "split" then set split cell color of o to v
	end tell
end setOneCellColor

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
