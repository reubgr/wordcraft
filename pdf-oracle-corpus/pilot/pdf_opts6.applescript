-- pdf_opts6.applescript — subcommand helper for the v6 pilot driver (#6153).
-- v6 fix vs v5: the accumulator `lines` and the parameter `line` collide
-- with the Word dictionary's line class/plural at runtime (-10006
-- "Can't set every line", -1728 "Can't get every text item of line") —
-- renamed to outLines / snapLine. Everything else is v5 verbatim.
-- ALL Track Changes Options access goes through the application `settings`
-- property (wSet) — the route proven by probe 3b (pdf_settings_form).
-- The bare `Word options` class-name reference is never used: it fails
-- with -1728 (get) / -10006 (set) on Word for Mac 16.114.4.
-- Subcommands:
--   snapshot                  -> one equality-based options snapshot line
--                                (per-field tries; a failing field records
--                                ERR<num> instead of failing the snapshot)
--   set-custom                -> isolated custom sets + SET lines + snapshot
--   restore <snapshot line>   -> apply snapshot values + snapshot line
--   export <mode> <docxPath> <containerPdf> <outPdf>
--                             -> open staged copy, set doc/view state,
--                                save as PDF to both paths, close without
--                                saving (saved-flag fallback for -1708)
--   view / set-view <name>    -> view revisions mode read / set
-- All enum literals live inside tell blocks (-2741 compile rule).
on run argv
	set cmd to item 1 of argv
	if cmd is "snapshot" then
		return my snapshotOptions("SNAP")
	else if cmd is "set-custom" then
		set out to {}
		try
			tell application "Microsoft Word" to set inserted text color of settings to blue
			set end of out to "SET_OK insColor=blue"
		on error errMsg number errNum
			set end of out to ("SET_FAIL insColor " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set inserted text mark of settings to inserted text mark underline
			set end of out to "SET_OK insMark=underline"
		on error errMsg number errNum
			set end of out to ("SET_FAIL insMark " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set deleted text color of settings to red
			set end of out to "SET_OK delColor=red"
		on error errMsg number errNum
			set end of out to ("SET_FAIL delColor " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to set deleted text mark of settings to deleted text mark strike through
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
		return my doExport(item 2 of argv, item 3 of argv, item 4 of argv, item 5 of argv)
	else if cmd is "set-view" then
		my setViewMode(item 2 of argv)
		return "VIEW_NOW " & my viewModeName()
	else if cmd is "view" then
		return "VIEW_MODE " & my viewModeName()
	end if
	return "UNKNOWN_CMD"
end run

-- Opens the staged copy, applies the mode's doc/view state, saves the PDF
-- to the container path and then to the out-dir path, and closes WITHOUT
-- saving the document. Close falls back to the showin pattern (mark the
-- document saved, close with no saving argument) when plain
-- `close saving no` errors (the -1708 wedge seen after save-as in probe 3b).
on doExport(mode, docxPath, containerPath, outPath)
	set outLines to {}
	with timeout of 150 seconds
		try
			tell application "Microsoft Word"
				open docxPath
				set doc to active document
				set end of outLines to "DOC_NAME " & (name of doc)
			end tell
		on error errMsg number errNum
			set end of outLines to ("OPEN_FAIL " & errNum & " " & errMsg)
			set AppleScript's text item delimiters to linefeed
			return outLines as text
		end try
		try
			set end of outLines to "VIEW_BEFORE " & my viewModeName()
		on error errMsg number errNum
			set end of outLines to ("VIEW_BEFORE_FAIL " & errNum)
		end try
		try
			tell application "Microsoft Word"
				set doc to active document
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
				set end of outLines to ("VIEW_STATE mode=" & mode & " showRev=" & (show revisions of doc as text) & " printRev=" & (print revisions of doc as text) & " revView=" & rv & " revMode=" & my viewModeName())
			end tell
		on error errMsg number errNum
			set end of outLines to ("EXPORT_SETUP_FAIL " & errNum & " " & errMsg)
			set AppleScript's text item delimiters to linefeed
			return outLines as text
		end try
		try
			tell application "Microsoft Word" to save as active document file name containerPath file format format PDF
			set end of outLines to "SAVE_CONTAINER_OK"
		on error errMsg number errNum
			set end of outLines to ("SAVE_CONTAINER_FAIL " & errNum & " " & errMsg)
		end try
		try
			tell application "Microsoft Word" to save as active document file name outPath file format format PDF
			set end of outLines to "SAVE_OUT_OK"
		on error errMsg number errNum
			set end of outLines to ("SAVE_OUT_FAIL " & errNum & " " & errMsg)
		end try
		-- close without saving; saved-flag fallback for the -1708 wedge
		try
			tell application "Microsoft Word" to close active document saving no
			set end of outLines to "CLOSE_OK plain"
		on error errMsg number errNum
			set end of outLines to ("CLOSE_PLAIN_FAIL " & errNum & " " & errMsg)
			try
				tell application "Microsoft Word"
					set saved of active document to true
					close active document
				end tell
				set end of outLines to "CLOSE_OK savedflag"
			on error errMsg2 number errNum2
				set end of outLines to ("CLOSE_FAIL " & errNum2 & " " & errMsg2)
			end try
		end try
		try
			tell application "Microsoft Word" to set end of outLines to ("FINAL_DOCS " & (count of documents))
		end try
	end timeout
	set AppleScript's text item delimiters to linefeed
	return outLines as text
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

-- Equality-based snapshot of the Track Changes Options via `settings`.
-- Each field is read in its own try: a field that errors records ERR<num>
-- and the snapshot still completes (the FORMAT.md rule: a check that
-- could not run must not report as though it did — and must not take the
-- other fields down with it).
on snapshotOptions(tag)
	set s to tag
	set s to s & tab & "insColor=" & my snapColor("ins")
	set s to s & tab & "insMark=" & my snapInsMark()
	set s to s & tab & "delColor=" & my snapColor("del")
	set s to s & tab & "delMark=" & my snapDelMark()
	set s to s & tab & "insCellColor=" & my snapCellColor("ins")
	set s to s & tab & "delCellColor=" & my snapCellColor("del")
	set s to s & tab & "mergedCellColor=" & my snapCellColor("merged")
	set s to s & tab & "splitCellColor=" & my snapCellColor("split")
	set s to s & tab & "moveFromColor=" & my snapColor("moveFrom")
	set s to s & tab & "moveFromMark=" & my snapMoveFromMark()
	set s to s & tab & "moveToColor=" & my snapColor("moveTo")
	set s to s & tab & "moveToMark=" & my snapMoveToMark()
	set s to s & tab & "commentsColor=" & my snapColor("comments")
	set s to s & tab & "printComments=" & my snapPrintComments()
	return s
end snapshotOptions

on snapColor(which)
	try
		tell application "Microsoft Word"
			set o to settings
			if which is "ins" then set v to inserted text color of o
			if which is "del" then set v to deleted text color of o
			if which is "moveFrom" then set v to move from text color of o
			if which is "moveTo" then set v to move to text color of o
			if which is "comments" then set v to comments color of o
		end tell
		return my colorName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapColor

on snapCellColor(which)
	try
		tell application "Microsoft Word"
			set o to settings
			if which is "ins" then set v to inserted cell color of o
			if which is "del" then set v to deleted cell color of o
			if which is "merged" then set v to merged cell color of o
			if which is "split" then set v to split cell color of o
		end tell
		return my cellColorName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapCellColor

on snapInsMark()
	try
		tell application "Microsoft Word" to set v to inserted text mark of settings
		return my insMarkName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapInsMark

on snapDelMark()
	try
		tell application "Microsoft Word" to set v to deleted text mark of settings
		return my delMarkName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapDelMark

on snapMoveFromMark()
	try
		tell application "Microsoft Word" to set v to move from text mark of settings
		return my moveFromMarkName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapMoveFromMark

on snapMoveToMark()
	try
		tell application "Microsoft Word" to set v to move to text mark of settings
		return my moveToMarkName(v)
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapMoveToMark

on snapPrintComments()
	try
		tell application "Microsoft Word" to return (print comments of settings) as text
	on error errMsg number errNum
		return "ERR" & errNum
	end try
end snapPrintComments

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

on valOf(snapLine, key)
	set AppleScript's text item delimiters to tab
	set parts to text items of snapLine
	set AppleScript's text item delimiters to ""
	repeat with p in parts
		set ps to p as text
		if ps starts with (key & "=") then return text ((length of key) + 2) thru -1 of ps
	end repeat
	return ""
end valOf

on restoreFromSnapshot(snapLine)
	set ic to my valOf(snapLine, "insColor")
	set im to my valOf(snapLine, "insMark")
	set dc to my valOf(snapLine, "delColor")
	set dm to my valOf(snapLine, "delMark")
	set icc to my valOf(snapLine, "insCellColor")
	set dcc to my valOf(snapLine, "delCellColor")
	set mcc to my valOf(snapLine, "mergedCellColor")
	set scc to my valOf(snapLine, "splitCellColor")
	set mfc to my valOf(snapLine, "moveFromColor")
	set mfm to my valOf(snapLine, "moveFromMark")
	set mtc to my valOf(snapLine, "moveToColor")
	set mtm to my valOf(snapLine, "moveToMark")
	set cc to my valOf(snapLine, "commentsColor")
	my applyAll(ic, im, dc, dm, icc, dcc, mcc, scc, mfc, mfm, mtc, mtm, cc)
end restoreFromSnapshot

-- Applies one snapshot's values through `settings`. Fields whose snapshot
-- value is empty or an ERR marker are skipped (nothing sane to restore to;
-- the RESTORED snapshot readback in the driver log shows the outcome).
on applyAll(ic, im, dc, dm, icc, dcc, mcc, scc, mfc, mfm, mtc, mtm, cc)
	tell application "Microsoft Word"
		set o to settings
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
	if nm is "" then return
	if nm starts with "ERR" then return
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
		set o to settings
		if which is "ins" then set inserted text color of o to v
		if which is "del" then set deleted text color of o to v
		if which is "moveFrom" then set move from text color of o to v
		if which is "moveTo" then set move to text color of o to v
		if which is "comments" then set comments color of o to v
	end tell
end setOneColor

on setCellColorProp(which, nm)
	if nm is "" then return
	if nm starts with "ERR" then return
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
		set o to settings
		if which is "ins" then set inserted cell color of o to v
		if which is "del" then set deleted cell color of o to v
		if which is "merged" then set merged cell color of o to v
		if which is "split" then set split cell color of o to v
	end tell
end setOneCellColor
