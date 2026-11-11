-- pdf_opts4.applescript — subcommand helper for the v4 pilot driver (#6153).
-- The bash driver orchestrates: it opens documents via LaunchServices
-- (grant-free) and calls this script per step, capturing stdout.
-- Subcommands:
--   snapshot                 -> one equality-based options snapshot line
--   set-custom               -> isolated custom sets + SET lines + snapshot
--   restore <snapshot line>  -> apply snapshot values + snapshot line
--   export <mode>            -> per-mode doc/view settings on the ACTIVE
--                               document, view snapshot line, VIEW_BEFORE
--   set-view <balloon|inline|mixed> -> set view revisions mode, print it
--   view                     -> print current view revisions mode name
-- All enum literals live inside tell blocks (-2741 compile rule).
on run argv
	set cmd to item 1 of argv
	if cmd is "snapshot" then
		return my snapshotOptions("SNAP")
	else if cmd is "set-custom" then
		set out to {}
		try
			tell application "Microsoft Word" to set inserted text color of Word options to blue
			set end of out to "SET_OK insColor=blue"
		on error errMsg number errNum
			set end of out to ("SET_FAIL insColor " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set inserted text mark of Word options to inserted text mark underline
			set end of out to "SET_OK insMark=underline"
		on error errMsg number errNum
			set end of out to ("SET_FAIL insMark " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set deleted text color of Word options to red
			set end of out to "SET_OK delColor=red"
		on error errMsg number errNum
			set end of out to ("SET_FAIL delColor " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set deleted text mark of Word options to deleted text mark strike through
			set end of out to "SET_OK delMark=strikethrough"
		on error errMsg number errNum
			set end of out to ("SET_FAIL delMark " & errNum & " " & errMsg)
		end try
		set end of out to my snapshotOptions("CUSTOM")
		set AppleScript's text item delimiters to linefeed
		return out as text
	else if cmd is "restore" then
		my restoreFromSnapshot(item 2 of argv)
		return my snapshotOptions("RESTORED")
	else if cmd is "export" then
		return my doExport(item 2 of argv, item 3 of argv, item 4 of argv)
	else if cmd is "set-view" then
		my setViewMode(item 2 of argv)
		return "VIEW_NOW " & my viewModeName()
	else if cmd is "view" then
		return "VIEW_MODE " & my viewModeName()
	end if
	return "UNKNOWN_CMD"
end run

-- Sets doc/view state for the mode on the active document and saves the
-- PDF twice: to the container path and to the out-dir path, each logged.
on doExport(mode, containerPath, outPath)
	set lines to {}
	try
		tell application "Microsoft Word"
			set doc to active document
			set end of lines to "DOC_NAME " & (name of doc)
			set end of lines to "VIEW_BEFORE " & my viewModeName()
			if mode is "markup-custom" then set revisions mode of view of active window to in line revisions
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
			set rv to "?"
			if (revisions view of view of active window) = revisions view final then set rv to "final"
			if (revisions view of view of active window) = revisions view original then set rv to "original"
			set end of lines to ("VIEW_STATE mode=" & mode & " showRev=" & (show revisions of doc as text) & " printRev=" & (print revisions of doc as text) & " revView=" & rv & " revMode=" & my viewModeName())
		end tell
	on error errMsg number errNum
		set end of lines to ("EXPORT_SETUP_FAIL " & errNum & " " & errMsg)
		set AppleScript's text item delimiters to linefeed
		return lines as text
	end try
	try
		tell application "Microsoft Word" to save as active document file name containerPath file format format PDF
		set end of lines to "SAVE_CONTAINER_OK"
	on error errMsg number errNum
		set end of lines to ("SAVE_CONTAINER_FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to save as active document file name outPath file format format PDF
		set end of lines to "SAVE_OUT_OK"
	on error errMsg number errNum
		set end of lines to ("SAVE_OUT_FAIL " & errNum & " " & errMsg)
	end try
	set AppleScript's text item delimiters to linefeed
	return lines as text
end doExport

on viewModeName()
	tell application "Microsoft Word"
		set m to revisions mode of view of active window
		if m = balloon revisions then return "balloon"
		if m = in line revisions then return "inline"
		if m = mixed revisions then return "mixed"
		return "unknown"
	end tell
end viewModeName

on setViewMode(nm)
	tell application "Microsoft Word"
		if nm is "balloon" then set revisions mode of view of active window to balloon revisions
		if nm is "inline" then set revisions mode of view of active window to in line revisions
		if nm is "mixed" then set revisions mode of view of active window to mixed revisions
	end tell
end setViewMode

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

on restoreFromSnapshot(line)
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
