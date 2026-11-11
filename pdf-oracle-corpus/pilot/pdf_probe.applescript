-- pdf_probe.applescript — recovery + Word options probe (runs under the
-- pdf-corpus-pilot lock claim left held by run 3). Closes every open
-- document WITHOUT saving (they are autorecovery/staged copies in a
-- restarted Word; originals live outside Word), then probes how the
-- Track Changes Options colour/mark properties behave through the
-- dictionary: equality-based readback (the `as text` coercion of the
-- colour property fails with -1700) and set attempts with verbatim
-- error capture. Logs to pdf_probe_log.txt in the container.
with timeout of 300 seconds
	set C to "/Users/muse/Library/Containers/com.microsoft.Word/Data/Documents/"
	set logPath to C & "pdf_probe_log.txt"
	my logLine(logPath, "PROBE_START")
	tell application "Microsoft Word"
		set n to count of documents
		my logLine(logPath, "DOCS_BEFORE " & n)
		set i to 1
		repeat while i ≤ (count of documents)
			try
				my logLine(logPath, "DOC " & i & " name=" & (name of document i))
			on error errMsg number errNum
				my logLine(logPath, "DOC " & i & " NAME_ERR " & errNum & " " & errMsg)
			end try
			set i to i + 1
		end repeat
		repeat while (count of documents) > 0
			close document 1 saving no
		end repeat
		my logLine(logPath, "DOCS_AFTER " & (count of documents))
	end tell
	-- readback by equality against candidate constants
	my logLine(logPath, "STATE_BEFORE " & my optState())
	-- set attempts, each isolated
	try
		tell application "Microsoft Word" to set inserted text mark of Word options to inserted text mark underline
		my logLine(logPath, "SET insMark=underline OK")
	on error errMsg number errNum
		my logLine(logPath, "SET insMark FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set inserted text color of Word options to blue
		my logLine(logPath, "SET insColor=blue OK")
	on error errMsg number errNum
		my logLine(logPath, "SET insColor FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set deleted text color of Word options to red
		my logLine(logPath, "SET delColor=red OK")
	on error errMsg number errNum
		my logLine(logPath, "SET delColor FAIL " & errNum & " " & errMsg)
	end try
	try
		tell application "Microsoft Word" to set deleted text mark of Word options to deleted text mark strike through
		my logLine(logPath, "SET delMark=strikethrough OK")
	on error errMsg number errNum
		my logLine(logPath, "SET delMark FAIL " & errNum & " " & errMsg)
	end try
	my logLine(logPath, "STATE_AFTER_SETS " & my optState())
	my logLine(logPath, "PROBE_DONE")
end timeout

on optState()
	try
		tell application "Microsoft Word"
			set o to Word options
			set ic to "?"
			if (inserted text color of o) = blue then set ic to "blue"
			if (inserted text color of o) = red then set ic to "red"
			if (inserted text color of o) = by author then set ic to "byAuthor"
			if (inserted text color of o) = auto then set ic to "auto"
			if (inserted text color of o) = black then set ic to "black"
			if (inserted text color of o) = pink then set ic to "pink"
			if (inserted text color of o) = green then set ic to "green"
			if (inserted text color of o) = violet then set ic to "violet"
			set dc to "?"
			if (deleted text color of o) = blue then set dc to "blue"
			if (deleted text color of o) = red then set dc to "red"
			if (deleted text color of o) = by author then set dc to "byAuthor"
			if (deleted text color of o) = auto then set dc to "auto"
			if (deleted text color of o) = black then set dc to "black"
			if (deleted text color of o) = pink then set dc to "pink"
			set im to "?"
			if (inserted text mark of o) = inserted text mark underline then set im to "underline"
			if (inserted text mark of o) = inserted text mark none then set im to "none"
			if (inserted text mark of o) = inserted text mark bold then set im to "bold"
			if (inserted text mark of o) = inserted text mark italic then set im to "italic"
			if (inserted text mark of o) = inserted text mark color only then set im to "colorOnly"
			set dm to "?"
			if (deleted text mark of o) = deleted text mark strike through then set dm to "strikethrough"
			if (deleted text mark of o) = deleted text mark none then set dm to "none"
			if (deleted text mark of o) = deleted text mark hidden then set dm to "hidden"
			if (deleted text mark of o) = deleted text mark caret then set dm to "caret"
			return "insColor=" & ic & " insMark=" & im & " delColor=" & dc & " delMark=" & dm & " printComments=" & (print comments of o as text)
		end tell
	on error errMsg number errNum
		return "OPTSTATE_FAIL " & errNum & " " & errMsg
	end try
end optState

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
